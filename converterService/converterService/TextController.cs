using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using Range = Microsoft.Office.Interop.Excel.Range;

[Route("text")]
[ApiController]
public class TextController : ControllerBase
{
    public static string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log.txt");

    [HttpPost]
    public async Task<IActionResult> ConvertToText(IFormFile file)
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
        if (format == "word") { ConvertWordToText(filePath, pdfPath); convertAgain = false; }
        else if (format == "pdf") { ConvertPdfToText(filePath, pdfPath); convertAgain = false; }
        else if (format == "excel") { ConvertExcelToText(filePath, pdfPath); convertAgain = false; }
        else if (format == "picture") ConvertPictureToPdf(filePath, pdfPath);
        else if (format == "powerPoint") ConvertPowerPointToPdf(filePath, pdfPath);
        else
        {
            WriteErrorMesage("Файл не может быть сконвертирован в PDF, так как формат не поддерживается");
            return BadRequest(new CustomError("Неверный формат", "Файл не может быть сконвертирован в PDF, так как формат не поддерживается", 400));
        }
        byte[] textBytes;
        if (convertAgain)
        {
            ConvertPdfToText(pdfPath, filePath);
            textBytes = await System.IO.File.ReadAllBytesAsync(filePath);
        }
        else textBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
        System.IO.File.Delete(filePath);
        System.IO.File.Delete(pdfPath);
        WriteSuccessMesage($"Файл {fileName} успешно конвертирован по алгоритму {format}");
        return File(textBytes, "text/plain", $"{Path.GetFileNameWithoutExtension(fileName)}_converted.txt");
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


    private void ConvertWordToText(string inputFile, string outputFile)
    {
        var wordApp = new Microsoft.Office.Interop.Word.Application();
        wordApp.Visible = false;
        wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
        var doc = wordApp.Documents.Open(inputFile);
        try
        {
            wordApp.Run("defaultMacro");
        }
        catch (Exception e) { }
   
        string textContent = doc.Content.Text;
        System.IO.File.WriteAllText(outputFile, textContent);

        doc.Close();
        wordApp.Quit();
    }

    private void ConvertPdfToText(string sourceFilePath, string destinationFilePath)
    {
        string ghostscriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ghostscript\\gswin64.exe");
        string deviceName = "txtwrite";
        string outputFileType = "-sOutputFile=";
        string arguments = $"-dNOPAUSE -dBATCH -sDEVICE={deviceName} {outputFileType}\"{destinationFilePath}\" \"{sourceFilePath}\"";

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

    private void ConvertExcelToText(string inputFile, string outputFile)
    {
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        Workbook workbook = excelApp.Workbooks.Open(inputFile);

        try
        {
            excelApp.Run("defaultMacro");
        }
        catch (Exception e) { }

        using (StreamWriter writer = new StreamWriter(outputFile))
        {
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                foreach (Range cell in worksheet.UsedRange)
                {
                    string cellText = cell.Value != null ? cell.Value.ToString() : "";
                    writer.WriteLine(cellText);
                }
            }
        }

        workbook.Close(false);
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
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