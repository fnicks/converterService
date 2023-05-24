using System.IO;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;

[Route("convert")]
[ApiController]
public class OfficeToPdfController : ControllerBase
{
    [HttpPost]
    public async Task<IActionResult> ConvertToPdf(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("Не найден файл в запросе");
        }

        var fileName = file.FileName;
        var filePath = Path.GetTempFileName();
        var pdfPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(fileName) + ".pdf");

        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        string format = checkFileType(fileName);
        if (format == "word") ConvertWordToPdf(filePath, pdfPath);
        else if (format == "excel") ConvertExcelToPdf(filePath, pdfPath);
        else if (format == "picture") ConvertPictureToPdf(filePath, pdfPath);
        else if (format == "powerPoint") ConvertPowerPointToPdf(filePath, pdfPath);
        else return BadRequest("Неверный формат");
        var pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
        System.IO.File.Delete(filePath);
        System.IO.File.Delete(pdfPath);

        return File(pdfBytes, "application/pdf", $"{Path.GetFileNameWithoutExtension(fileName)}.pdf");
    }

    private void ConvertWordToPdf(string inputFile, string outputFile)
    {
        var wordApp = new Microsoft.Office.Interop.Word.Application();
        wordApp.Visible = false;
        wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
        var doc = wordApp.Documents.Open(inputFile);
        try
        {
            wordApp.Run("defaultMacro");
        } catch (Exception e) { }
        doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF, UseISO19005_1: true);
        doc.Close();
        wordApp.Quit();
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
        string[] word = { "doc", "docx", "docm", "rtf", "xml", "pdf", "odt", "txt" };
        string[] powerPoint = { "pptx", "pptm", "ppt" };
        string[] picture = { "jpg", "jpeg", "png" };
        string[] excel = { "xls", "xlsx", "csv" };
        string[] array = fileName.Split(".");
        int lenght = array.Length;
        if (lenght > 0)
        {
            string extension = array[lenght - 1];
            if (word.Contains(extension)) return "word";
            if (excel.Contains(extension)) return "excel";
            if (picture.Contains(extension)) return "picture";
            if (powerPoint.Contains(extension)) return "powerPoint";

            return "Error";
        }
        else return "Error";
    }
}

