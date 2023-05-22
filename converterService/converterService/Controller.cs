using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

[Route("convert")]
[ApiController]
public class OfficeToPdfController : ControllerBase
{
    [HttpPost]
    public async Task<IActionResult> ConvertToPdf(IFormFile officeFile)
    {
        if (officeFile == null || officeFile.Length == 0)
        {
            return BadRequest("Не найден файл в запросе");
        }

        var fileName = officeFile.FileName;
        var filePath = Path.GetTempFileName();
        var pdfPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(fileName) + ".pdf");

        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await officeFile.CopyToAsync(stream);
        }

        string format = checkFileType(fileName);
        if (format == "word") ConvertWordToPdf(filePath, pdfPath);
        else if (format == "excel") ConvertExcelToPdf(filePath, pdfPath);
        else return BadRequest("Неверный формат");


        var pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
        System.IO.File.Delete(filePath);
        System.IO.File.Delete(pdfPath);

        return File(pdfBytes, "application/pdf", $"{fileName}.pdf");
    }

    private void ConvertWordToPdf(string inputFile, string outputFile)
    {
        var wordApp = new Microsoft.Office.Interop.Word.Application();
        wordApp.Visible = false;
        wordApp.DisplayAlerts = 0;
        var doc = wordApp.Documents.Open(inputFile);
        try
        {
            wordApp.Run("defaultMacro");
        } catch (Exception e) { }
        doc.SaveAs2(outputFile, WdSaveFormat.wdFormatPDF);
        doc.Close();
        wordApp.Quit();
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
        string[] word = { "doc", "docx", "docm", "rtf", "xml" };
        string[] excel = { "xls", "xlsx", "csv" };
        string[] array = fileName.Split(".");
        int lenght = array.Length;
        if (lenght > 0)
        {
            string extension = array[lenght - 1];
            if (word.Contains(extension)) return "word";
            if (excel.Contains(extension)) return "excel";
            return "Error";
        }
        else return "Error";
    }
}