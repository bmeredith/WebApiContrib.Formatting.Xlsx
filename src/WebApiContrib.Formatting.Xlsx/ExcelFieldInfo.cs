namespace WebApiContrib.Formatting.Xlsx
{
    public class ExcelFieldInfo
    {
        public string PropertyName { get; set; }
        public ExcelColumnAttribute ExcelAttribute { get; set; }
        public string FormatString { get; set; }
        public string Header { get; set; }

        public string ExcelNumberFormat => ExcelAttribute?.NumberFormat;

        public bool IsExcelHeaderDefined => ExcelAttribute?.Header != null;

        public ExcelFieldInfo(string propertyName, ExcelColumnAttribute excelAttribute = null, string formatString = null)
        {
            PropertyName = propertyName;
            ExcelAttribute = excelAttribute;
            FormatString = formatString;
            Header = IsExcelHeaderDefined ? ExcelAttribute?.Header : propertyName;
        }
    }
}
