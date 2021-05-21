using App.Programs;
using Others;

namespace App
{
    public static class Program
    {

        private const string excelFilePath = @"D:\DatosPruebas\Tuya\4-CertiReten\Certificados\2020\C1128400420.xlsx";
        private const string excelDataSettings  = @"D:\0-Clientes\2-Tuya\2-Asignaciones\6-CertificadosRetenExcelToPdf\2-AlternativaPlantilla\excelDataConfig.json";

        static void Main(string[] args)
        {

            //Programs

            //ExcelConverter.ConverterToPdf();
            string pathFile = ExcelReader.GetCertificatePDF(excelFilePath, excelDataSettings, string.Empty);
        }
    }
}
