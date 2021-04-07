using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelToPdf
{
    public class ExcelApplicationWrapper : IDisposable
    {
        public Application ExcelApplication { get; }

        public ExcelApplicationWrapper()
        {
            ExcelApplication = new Application();
        }

        public void Dispose()
        {
            // Each file I open is locked by the background EXCEL.exe until it is quitted
            ExcelApplication.Quit();
            Marshal.ReleaseComObject(ExcelApplication);
        }
    }

    public class MicrosoftOfficeInteropExcel : IConvertExcelToPdf
    {
        public void ConvertExcelToPdf(string originFilePath, string destinationPath)
        {
            using (var excelApplication = new ExcelApplicationWrapper())
            {
                var thisFileWorkbook = excelApplication.ExcelApplication.Workbooks.Open(originFilePath);
                string newPdfFilePath = Path.Combine(destinationPath, $"MicrosoftOfficeInteropExcel_{Path.GetFileNameWithoutExtension(originFilePath)}.pdf");
                                
                thisFileWorkbook.ExportAsFixedFormat(
                    XlFixedFormatType.xlTypePDF,
                    newPdfFilePath);

                thisFileWorkbook.Close(false, originFilePath);
                Marshal.ReleaseComObject(thisFileWorkbook);
            }
        }
    }
}
