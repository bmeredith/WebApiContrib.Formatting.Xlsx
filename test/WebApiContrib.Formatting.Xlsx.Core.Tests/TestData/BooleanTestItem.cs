namespace WebApiContrib.Formatting.Xlsx.Core.Tests.TestData
{
    public class BooleanTestItem
    {
        public bool Value1 { get; set; }

        [ExcelColumn(TrueValue="Yes", FalseValue="No")]
        public bool Value2 { get; set; }

        public bool? Value3 { get; set; }

        [ExcelColumn(TrueValue = "Yes", FalseValue = "No")]
        public bool? Value4 { get; set; }
    }
}
