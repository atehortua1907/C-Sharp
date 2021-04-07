using System.IO;

namespace ExcelToPdf
{
    public class Sautinsoft : IConvertExcelToPdf
    {
        public void ConvertExcelToPdf(string originPath, string destinationPath)
        {
            SautinSoft.ExcelToPdf x = new SautinSoft.ExcelToPdf();
            string outputFilePath = Path.Combine(destinationPath, $"Sautinsoft_{Path.GetFileNameWithoutExtension(originPath)}.pdf");
            x.ConvertFile(originPath, outputFilePath);
        }
    }
}
