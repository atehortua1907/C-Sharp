using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace ExcelToPdf
{
    public class ItextSharp : IConvertExcelToPdf
    {
        public void ConvertExcelToPdf(string originPath, string destinationPath)
        {
            string outputFilePath = Path.Combine(destinationPath, $"itextSharp_{Path.GetFileNameWithoutExtension(originPath)}.pdf");

            using (StreamReader fileRead = new StreamReader(originPath))
            {                
                Document doc = new Document();
                PdfWriter.GetInstance(doc, new FileStream(outputFilePath, FileMode.Create));
                doc.Open();
                doc.Add(new Paragraph(fileRead.ReadToEnd()));
                doc.Close();
            }
        }
    }
}
