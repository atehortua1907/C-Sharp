using Spire.Xls;
using System.IO;

namespace ExcelToPdf
{
    public class SpireXls : IConvertExcelToPdf
    {
        public void ConvertExcelToPdf(string originFilePath, string destinationPath)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(originFilePath);
            string outputFilePath = Path.Combine(destinationPath, $"SpireXls_{Path.GetFileNameWithoutExtension(originFilePath)}.pdf");
            workbook.SaveToFile(outputFilePath, FileFormat.PDF);
        }
    }
}
