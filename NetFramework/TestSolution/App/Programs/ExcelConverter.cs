using ExcelToPdf;

namespace App.Programs
{
    public static class ExcelConverter
    {
        private const string OriginFilePath = @"D:\1-Repositorios\3-David\C-Sharp\NetFramework\TestSolution\ExcelToPdf\Archivos\EjemploFactura.xlsx";
        private const string DestinationPath = @"D:\1-Repositorios\3-David\C-Sharp\NetFramework\TestSolution\ExcelToPdf\Archivos";

        public static void ConverterToPdf()
        {
            IConvertExcelToPdf convertExcelToPdf = ConvertersFactory.GetConverter(ConverterType.SyncfusionExcelToPdfConverter);
            convertExcelToPdf.ConvertExcelToPdf(OriginFilePath, DestinationPath);
        }
    }
}
