using ExcelDataReader;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Others
{
    public static class ExcelReader
    {

        public static string GetCertificatePDF(string excelFilePath, string excelDataSettings, string pathTempDownload)
        {            
            ExcelDataConfig excelDataConfig = GetExcelRead(excelFilePath, excelDataSettings);
            string newPdfFilePath = GetPathPdfFile(excelDataConfig, pathTempDownload);
            return newPdfFilePath;
        }

        public static ExcelDataConfig GetExcelRead(string filePath, string excelDataSettings)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTable table = result.Tables[0];
                    DataRowCollection data = table.Rows[0].Table.Rows;

                    string jsonConfigurations = string.Empty;
                    using (StreamReader sr = new StreamReader(excelDataSettings))
                        jsonConfigurations = sr.ReadToEnd();

                    ExcelDataConfig excelDataConfig = JsonConvert.DeserializeObject<ExcelDataConfig>(jsonConfigurations);
                    excelDataConfig.GetType().GetProperties().ToList().ForEach(propertyInfo => 
                    {
                        var property = propertyInfo.GetValue(excelDataConfig) as DataConfig;
                        property.Value = data[property.Row][property.Column].ToString();
                    });

                    return excelDataConfig;
                }
            }
        }

        private static string GetPathPdfFile(ExcelDataConfig excelDataConfig, string pathTempDownload)
        {
            string pdfTemplate = @"D:\0-Clientes\2-Tuya\2-Asignaciones\6-CertificadosRetenExcelToPdf\2-AlternativaPlantilla\PdfPlantillaCertificados.pdf";
            string pdfResult = @"D:\0-Clientes\2-Tuya\2-Asignaciones\6-CertificadosRetenExcelToPdf\2-AlternativaPlantilla\pdfResult.pdf";
            PdfReader pdfReader = new PdfReader(pdfTemplate);
            FileStream stream = new FileStream(pdfResult, FileMode.Create);
            PdfStamper pdfStamper = new PdfStamper(pdfReader, stream);
            excelDataConfig.GetType().GetProperties().ToList().ForEach(propertyInfo =>
            {
                var property = propertyInfo.GetValue(excelDataConfig) as DataConfig;
                pdfStamper.AcroFields.SetField(propertyInfo.Name, property.Value);
            });
            pdfStamper.FormFlattening = true;
            pdfStamper.Close();
            return pdfResult;
        }
    }
}
