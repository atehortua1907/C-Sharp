namespace ExcelToPdf
{
    public enum ConverterType
    {
        ItexSharp = 1,
        MicrosoftOfficeInteropExcel = 2,
        SpireXls = 3,
        SyncfusionExcelToPdfConverter = 4,
        GroupDocs = 5,
        Sautinsoft = 6
    }

    public static class ConvertersFactory
    {
        public static IConvertExcelToPdf GetConverter(ConverterType converterType)
        {
            IConvertExcelToPdf convertExcelToPdf;
            switch (converterType)
            {
                case ConverterType.ItexSharp:
                    convertExcelToPdf = new ItextSharp();
                    break;
                case ConverterType.MicrosoftOfficeInteropExcel:                    
                    convertExcelToPdf =  new MicrosoftOfficeInteropExcel();
                    break;
                case ConverterType.SpireXls:
                    convertExcelToPdf = new SpireXls();
                    break;
                case ConverterType.SyncfusionExcelToPdfConverter:
                    convertExcelToPdf = new SyncfusionExcelToPdfConverter();
                    break;
                case ConverterType.GroupDocs:
                    convertExcelToPdf = new GroupDocs();
                    break;
                case ConverterType.Sautinsoft:
                    convertExcelToPdf = new Sautinsoft();
                    break;
                default:
                    convertExcelToPdf = new SyncfusionExcelToPdfConverter();
                    break;
            }

            return convertExcelToPdf;
        }
    }
}
