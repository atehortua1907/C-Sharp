using Syncfusion.ExcelToPdfConverter;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using System.IO;

namespace ExcelToPdf
{
    public class SyncfusionExcelToPdfConverter : IConvertExcelToPdf
    {
        public void ConvertExcelToPdf(string originFilePath, string destinationPath)
        {
            using(ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;

                IWorkbook workbook = application.Workbooks.Open(originFilePath, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //convert the sheet to PDF
                ExcelToPdfConverter converter = new ExcelToPdfConverter(sheet);

                PdfDocument pdfDocument = new PdfDocument();
                pdfDocument = converter.Convert();
                string outputFilePath = Path.Combine(destinationPath, $"SyncfusionExcelToPdfConverter_{Path.GetFileNameWithoutExtension(originFilePath)}.pdf");
                pdfDocument.Save(outputFilePath);
            }
        }
    }
}
