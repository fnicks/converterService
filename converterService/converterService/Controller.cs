using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

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
    [HttpPost]
    public async Task<IActionResult> ConvertToPdf(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest(new CustomError("Отсутвует файл в запросе", "Файл должен называется file", 400));
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
        else return BadRequest(new CustomError("Неверный формат", "Файл не может быть сконвертирован в PDF, так как формат не поддерживается", 400));
        byte[] pdfBytes;
        if (convertAgain)
        {
            ConvertToPdfA(pdfPath, filePath);
            pdfBytes = await System.IO.File.ReadAllBytesAsync(filePath);
        }
        else pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
        System.IO.File.Delete(filePath);
        System.IO.File.Delete(pdfPath);

        return File(pdfBytes, "application/pdf", $"{Path.GetFileNameWithoutExtension(fileName)}_converted.pdf");
    }

    private void ConvertWordToPdf(string inputFile, string outputFile)
    {
        var wordApp = new Application();
        wordApp.Visible = true;
        wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
        var doc = wordApp.Documents.Open(inputFile);

        try
        {
            wordApp.Run("defaultMacro");
        } catch (Exception e) { }
        doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF);
        doc.Close();
        wordApp.Quit();
    }

    public void ConvertToPdfA(string sourceFilePath, string destinationFilePath)
    {
        string ghostscriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ghostscript\\gswin64c.exe");
        string deviceName = "pdfwrite";
        string outputFileType = "-sOutputFile=";
        string pdfaParams = "-dPDFA=1 -dPDFACompatibilityPolicy=1 -dPDFACompatibilityPolicy=1";
        string outputIntentParams = "-sColorConversionStrategy=UseDeviceIndependentColor";
        string arguments = $"-dNOPAUSE -dBATCH -dSAFER -sDEVICE={deviceName} {outputFileType}\"{destinationFilePath}\" {pdfaParams} {outputIntentParams} \"{sourceFilePath}\"";
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
        string[] word = { "doc", "docx", "docm", "rtf", "xml", "pdf", "odt", "txt", "wbk", "txt", ".docx", ".docm", ".dotx", ".dotm", ".docb" };
        string[] powerPoint = { "pptx", "pptm", "ppt" };
        string[] picture = { "jpg", "jpeg", "png", "tiff", "tif" };
        string[] excel = { "xls", "xlsx", "csv" };
        string[] array = fileName.Split(".");
        int lenght = array.Length;
        if (lenght > 0)
        {
            string extension = array[lenght - 1];
            if (word.Contains(extension)) return "word";
            else if (excel.Contains(extension)) return "excel";
            else if (picture.Contains(extension)) return "picture";
            else if (powerPoint.Contains(extension)) return "powerPoint";
            else if (extension == "pdf") return "pdf";
            else return "Error";
        }
        else return "Error";
    }
}

