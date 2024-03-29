﻿using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;

public class CustomError
{
    public string error { get; set; }
    public string description { get; set; }
    public int status { get; set; }

    public CustomError(string error, string description, int status)
    {
        this.error = error;
        this.description = description;
        this.status = status;
    }
}

[Route("convert")]
[ApiController]
public class OfficeToPdfController : ControllerBase
{
    public static string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log.txt");

    [HttpPost]
    public async Task<IActionResult> ConvertToPdf(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            WriteErrorMesage("Файл должен называется file");
            return BadRequest(new CustomError("Отсутвует файл в запросе", "Файл должен называется file", 400));
        }
        int convertAtOnce = 3;
        try
        {
            convertAtOnce = int.Parse(Environment.GetEnvironmentVariable("CONVERT_AT_ONCE"));
        }
        catch
        {
            WriteErrorMesage("Не получилось достать CONVERT_AT_ONCE из .env файла ");
        }

        Process[] ghostscriptProcesses = Process.GetProcessesByName("gswin64");
        if (ghostscriptProcesses.Length > convertAtOnce)
        {
            WriteErrorMesage($"Слишком много запросов. Одновременно может конвертироваться только {convertAtOnce} файлов");
            return new ObjectResult(new CustomError("Слишком много запросов", $"Одновременно может конвертироваться только {convertAtOnce} файлов", 429))
            {
                StatusCode = 429
            };
        }

        var fileName = file.FileName;
        var filePath = Path.GetTempFileName();
        var pdfPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(fileName) + "_converted.pdf");

        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }
        Boolean convertAgain = true;
        string format = checkFileType(fileName);
        if (format == "word") ConvertWordToPdf(filePath, pdfPath);
        else if (format == "pdf") { ConvertToPdfA(filePath, pdfPath); convertAgain = false; }
        else if (format == "excel") ConvertExcelToPdf(filePath, pdfPath);
        else if (format == "picture") ConvertPictureToPdf(filePath, pdfPath);
        else if (format == "powerPoint") ConvertPowerPointToPdf(filePath, pdfPath);
        else {
            WriteErrorMesage("Файл не может быть сконвертирован в PDF, так как формат не поддерживается");
            return BadRequest(new CustomError("Неверный формат", "Файл не может быть сконвертирован в PDF, так как формат не поддерживается", 400)); }
        byte[] pdfBytes;
        if (convertAgain)
        {
            ConvertToPdfA(pdfPath, filePath);
            pdfBytes = await System.IO.File.ReadAllBytesAsync(filePath);
        }
        else pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
        System.IO.File.Delete(filePath);
        System.IO.File.Delete(pdfPath);
        WriteSuccessMesage($"Файл {fileName} успешно конвертирован по алгоритму {format}");
        return File(pdfBytes, "application/pdf", $"{Path.GetFileNameWithoutExtension(fileName)}_converted.pdf");
    }

    public static void WriteErrorMesage(string message)
    {
        Console.ForegroundColor = ConsoleColor.DarkRed;
        Console.WriteLine(message);
        Console.ResetColor();
        string formattedDateTime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
        System.IO.File.AppendAllText(logFilePath, formattedDateTime + " | ERROR: " + message + Environment.NewLine);
    }

    public static void WriteSuccessMesage(string message)
    {
        Console.ForegroundColor = ConsoleColor.DarkGreen;
        Console.WriteLine(message);
        Console.ResetColor();
        string formattedDateTime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
        System.IO.File.AppendAllText(logFilePath, formattedDateTime + " | SUCCESS: " + message + Environment.NewLine);
    }

    static void LoadEnvironmentVariablesFromFile()
    {
        string filePath = "./.env";
        if (System.IO.File.Exists(filePath))
        {
            foreach (string line in System.IO.File.ReadAllLines(filePath))
            {
                string[] parts = line.Split('=', 2);
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim();
                    string value = parts[1].Trim();
                    Environment.SetEnvironmentVariable(key, value);
                }
            }
        }
        else
        {
            WriteErrorMesage($"Не найден .env файл: {filePath}");
        }
    }


    private void ConvertWordToPdf(string inputFile, string outputFile)
    {
        var wordApp = new Word.Application();
        wordApp.Visible = false;
        wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
        var doc = wordApp.Documents.Open(inputFile);
        try
        {
            wordApp.Run("defaultMacro");
        }
        catch (Exception e) { }
        doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF);
        doc.Close();
        wordApp.Quit();
    }

    public void ConvertToPdfA(string sourceFilePath, string destinationFilePath)
    {
        string ghostscriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ghostscript\\gswin64.exe");
        string deviceName = "pdfwrite";
        string outputFileType = "-sOutputFile=";
        string pdfaParams = "-dPDFA=1 -dPDFACompatibilityPolicy=1 -dPDFACompatibilityPolicy=1";
        string outputIntentParams = "-sColorConversionStrategy=UseDeviceIndependentColor";
        string arguments = $"-dNOPAUSE -dFastWebView=true -dBATCH -sDEVICE={deviceName} {outputFileType}\"{destinationFilePath}\" {pdfaParams} {outputIntentParams} \"{sourceFilePath}\"";
        Process process = new Process();
        ProcessStartInfo startInfo = new ProcessStartInfo(ghostscriptPath, arguments);
        startInfo.RedirectStandardOutput = true;
        startInfo.UseShellExecute = false;
        startInfo.CreateNoWindow = true;
        process.StartInfo = startInfo;
        process.Start();
        process.WaitForExit();
    }

    private void ConvertPowerPointToPdf(string inputFile, string outputFile)
    {
        var powerPointApp = new Microsoft.Office.Interop.PowerPoint.Application();
        powerPointApp.DisplayAlerts = Microsoft.Office.Interop.PowerPoint.PpAlertLevel.ppAlertsNone;
        var doc = powerPointApp.Presentations.Open(inputFile, MsoTriState.msoFalse, MsoTriState.msoFalse,
    WithWindow: MsoTriState.msoFalse);
        try
        {
            powerPointApp.Run("defaultMacro");
        }
        catch (Exception e) { }
        doc.ExportAsFixedFormat(outputFile, Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF, UseISO19005_1: true);
        doc.Close();
        powerPointApp.Quit();
    }

    private void ConvertPictureToPdf(string inputFile, string outputFile)
    {
        DateTime now = DateTime.Now;
        string customFormat = now.ToString("yyyy.MM.dd.HH.mm.ss");
        var tempPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(customFormat) + ".docx");
        var wordApp = new Microsoft.Office.Interop.Word.Application();
        //  wordApp.Documents.Add(tempPath);
        wordApp.Visible = false;
        wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
        Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();
        doc.SaveAs2(tempPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
        Microsoft.Office.Interop.Word.Range drange = doc.Range();
        InlineShape picture = drange.InlineShapes.AddPicture(inputFile, Type.Missing, Type.Missing, Type.Missing);
        try
        {
            wordApp.Run("defaultMacro");
        }
        catch (Exception e) { }
        doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF, UseISO19005_1: true);
        doc.Close();
        wordApp.Quit();
        System.IO.File.Delete(tempPath);
    }

    private void ConvertExcelToPdf(string inputFile, string outputFile)
    {
        var excelApp = new Microsoft.Office.Interop.Excel.Application();
        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
        excelApp.AskToUpdateLinks = false;
        excelApp.AlertBeforeOverwriting = false;
        var doc = excelApp.Workbooks.Open(inputFile);
        try
        {
            excelApp.Run("defaultMacro");
        }
        catch (Exception e) { }
        doc.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputFile);
        doc.Close(false);
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
    }


    private string checkFileType(string fileName)
    {
        FileTypes fileTypes = readJson();
        string[] array = fileName.Split(".");
        int lenght = array.Length;
        if (lenght > 0)
        {
            string extension = array[lenght - 1];
            if (fileTypes.Word.Contains(extension)) return "word";
            else if (fileTypes.Excel.Contains(extension)) return "excel";
            else if (fileTypes.Picture.Contains(extension)) return "picture";
            else if (fileTypes.PowerPoint.Contains(extension)) return "powerPoint";
            else if (extension == "pdf") return "pdf";
            else return "Error";
        }
        else return "Error";
    }

    private FileTypes readJson()
    {
        string filePath = "./appsettings.json";
        if (System.IO.File.Exists(filePath))
        {
            string json = System.IO.File.ReadAllText(filePath);
            FileTypes jsonData = JsonConvert.DeserializeObject<FileTypes>(json);
            return jsonData;
        }
        else
        {
            Console.WriteLine($"Не найден файл настроек: {filePath}");
        }
        return new FileTypes();
    }
}


public class FileTypes
{
    public string[] Word { get; set; }
    public string[] PowerPoint { get; set; }
    public string[] Picture { get; set; }
    public string[] Excel { get; set; }
}
