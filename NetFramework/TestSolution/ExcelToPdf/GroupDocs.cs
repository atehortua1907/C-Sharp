using GroupDocs.Conversion;
using GroupDocs.Conversion.Options.Convert;
using System.IO;

namespace ExcelToPdf
{
    public class GroupDocs : IConvertExcelToPdf
    {
        public void ConvertExcelToPdf(string originPath, string destinationPath)
        {
            using (Converter converter = new Converter(originPath))
            {
                PdfConvertOptions options = new PdfConvertOptions();
                string outputFilePath = Path.Combine(destinationPath, $"GroupDocs_{Path.GetFileNameWithoutExtension(originPath)}.pdf");
                converter.Convert(outputFilePath, options);
            }
        }
    }
}
