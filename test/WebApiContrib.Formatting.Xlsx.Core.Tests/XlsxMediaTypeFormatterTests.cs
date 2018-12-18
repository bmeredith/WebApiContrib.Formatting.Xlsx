using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Formatters;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.AspNetCore.WebUtilities;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using WebApiContrib.Formatting.Xlsx.Core.Tests.Fakes;
using WebApiContrib.Formatting.Xlsx.Core.Tests.TestData;
using Xunit;

namespace WebApiContrib.Formatting.Xlsx.Core.Tests
{
    public class XlsxMediaTypeFormatterTests
    {
        private const string XlsMimeType = "application/vnd.ms-excel";
        private const string XlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        [Fact]
        public void SupportedMediaTypes_SupportsExcelMediaTypes()
        {
            var formatter = new XlsxMediaTypeFormatter();

            formatter.SupportedMediaTypes.Any(mediaType => mediaType == XlsMimeType).Should()
                .BeTrue("XLS media type not supported.");
            formatter.SupportedMediaTypes.Any(mediaType => mediaType == XlsxMimeType).Should()
                .BeTrue("XLS media type not supported.");
        }

        [Fact]
        public void WriteToStreamAsync_WithListOfSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" } }.ToList();

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(3, "Worksheet should have three rows (including header column).");
            sheet.Dimension.End.Column.Should().Be(2, "Worksheet should have two columns.");
            sheet.GetValue<string>(1, 1).Should().Be("Value1", "Header in A1 is incorrect.");
            sheet.GetValue<string>(1, 2).Should().Be("Value2", "Header in B1 is incorrect.");
            sheet.GetValue<string>(2, 1).Should().Be(data[0].Value1, "Value in A2 is incorrect.");
            sheet.GetValue<string>(2, 2).Should().Be(data[0].Value2, "Value in B2 is incorrect.");
            sheet.GetValue<string>(3, 1).Should().Be(data[1].Value1, "Value in A3 is incorrect.");
            sheet.GetValue<string>(3, 2).Should().Be(data[1].Value2, "Value in B3 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithArrayOfSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(3, "Worksheet should have three rows (including header column).");
            sheet.Dimension.End.Column.Should().Be(2, "Worksheet should have two columns.");
            sheet.GetValue<string>(1, 1).Should().Be("Value1", "Header in A1 is incorrect.");
            sheet.GetValue<string>(1, 2).Should().Be("Value2", "Header in B1 is incorrect.");
            sheet.GetValue<string>(2, 1).Should().Be(data[0].Value1, "Value in A2 is incorrect.");
            sheet.GetValue<string>(2, 2).Should().Be(data[0].Value2, "Value in B2 is incorrect.");
            sheet.GetValue<string>(3, 1).Should().Be(data[1].Value1, "Value in A3 is incorrect.");
            sheet.GetValue<string>(3, 2).Should().Be(data[1].Value2, "Value in B3 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithArrayOfFormatStringTestItem_ValuesFormattedAppropriately()
        {
            var tomorrow = DateTime.Today.AddDays(1);
            var formattedDate = tomorrow.ToString("D");

            var data = new[] { new FormatStringTestItem { Value1 = tomorrow,
                                                          Value2 = tomorrow,
                                                          Value3 = tomorrow,
                                                          Value4 = tomorrow },

                               new FormatStringTestItem { Value1 = tomorrow,
                                                          Value2 = null,
                                                          Value3 = null,
                                                          Value4 = tomorrow } };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(3, "Worksheet should have three rows (including header column).");
            sheet.Dimension.End.Column.Should().Be(4, "Worksheet should have four columns.");
            sheet.GetValue<string>(1, 1).Should().Be("Value1", "Header in A1 is incorrect.");
            sheet.GetValue<string>(1, 2).Should().Be("Value2", "Header in B1 is incorrect.");
            sheet.GetValue<string>(1, 3).Should().Be("Value3", "Header in C1 is incorrect.");
            sheet.GetValue<string>(1, 4).Should().Be("Value4", "Header in D1 is incorrect.");
            sheet.GetValue<string>(2, 1).Should().NotBe(formattedDate, "Value in A2 is incorrect.");
            sheet.GetValue<string>(2, 2).Should().Be(formattedDate, "Value in B2 is incorrect.");
            sheet.GetValue<string>(2, 3).Should().NotBe(formattedDate, "Value in C2 is incorrect.");
            sheet.GetValue<string>(2, 4).Should().NotBe(formattedDate, "Value in D2 is incorrect.");
            sheet.GetValue<string>(3, 1).Should().NotBe(formattedDate, "Value in A3 is incorrect.");
            sheet.GetValue<string>(3, 2).Should().Be(string.Empty, "Value in B3 is incorrect.");
            sheet.GetValue<string>(3, 3).Should().Be(string.Empty, "Value in C3 is incorrect.");
            sheet.GetValue<string>(3, 4).Should().NotBe(formattedDate, "Value in D3 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithArrayOfBooleanTestItem_TrueOrFalseValueUsedAsAppropriate()
        {
            var data = new[] { new BooleanTestItem { Value1 = true,
                                                     Value2 = true,
                                                     Value3 = true,
                                                     Value4 = true },

                               new BooleanTestItem { Value1 = false,
                                                     Value2 = false,
                                                     Value3 = false,
                                                     Value4 = false },

                               new BooleanTestItem { Value1 = true,
                                                     Value2 = true,
                                                     Value3 = null,
                                                     Value4 = null } };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(4, "Worksheet should have four rows (including header column).");
            sheet.Dimension.End.Column.Should().Be(4, "Worksheet should have four columns.");
            sheet.GetValue<string>(1, 1).Should().Be("Value1", "Header in A1 is incorrect.");
            sheet.GetValue<string>(1, 2).Should().Be("Value2", "Header in B1 is incorrect.");
            sheet.GetValue<string>(1, 3).Should().Be("Value3", "Header in C1 is incorrect.");
            sheet.GetValue<string>(1, 4).Should().Be("Value4", "Header in D1 is incorrect.");
            sheet.GetValue<string>(2, 1).Should().Be("True", "Value in A2 is incorrect.");
            sheet.GetValue<string>(2, 2).Should().Be("Yes", "Value in B2 is incorrect.");
            sheet.GetValue<string>(2, 3).Should().Be("True", "Value in C2 is incorrect.");
            sheet.GetValue<string>(2, 4).Should().Be("Yes", "Value in D2 is incorrect.");
            sheet.GetValue<string>(3, 1).Should().Be("False", "Value in A3 is incorrect.");
            sheet.GetValue<string>(3, 2).Should().Be("No", "Value in B3 is incorrect.");
            sheet.GetValue<string>(3, 3).Should().Be("False", "Value in C3 is incorrect.");
            sheet.GetValue<string>(3, 4).Should().Be("No", "Value in D3 is incorrect.");
            sheet.GetValue<string>(4, 1).Should().Be("True", "Value in A4 is incorrect.");
            sheet.GetValue<string>(4, 2).Should().Be("Yes", "Value in B4 is incorrect.");
            sheet.GetValue<string>(4, 3).Should().Be(string.Empty, "Value in C4 is incorrect.");
            sheet.GetValue<string>(4, 4).Should().Be(string.Empty, "Value in D4 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(2, "Worksheet should have two rows (including header column).");
            sheet.Dimension.End.Column.Should().Be(2, "Worksheet should have two columns.");
            sheet.GetValue<string>(1, 1).Should().Be("Value1", "Header in A1 is incorrect.");
            sheet.GetValue<string>(1, 2).Should().Be("Value2", "Header in B1 is incorrect.");
            sheet.GetValue<string>(2, 1).Should().Be(data.Value1, "Value in A2 is incorrect.");
            sheet.GetValue<string>(2, 2).Should().Be(data.Value2, "Value in B2 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithComplexTestItem_WritesExcelDocumentToStream()
        {
            var data = new ComplexTestItem
            {
                Value1 = "Item 1",
                Value2 = DateTime.Today,
                Value3 = true,
                Value4 = 100.1,
                Value5 = TestEnum.First,
                Value6 = "Ignored"
            };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(2, "Worksheet should have two rows (including header column).");
            sheet.Dimension.End.Column.Should().Be(5, "Worksheet should have five columns.");
            sheet.GetValue<string>(1, 1).Should().Be("Header 4", "Header in A1 is incorrect.");
            sheet.GetValue<string>(1, 2).Should().Be("Value1", "Header in B1 is incorrect.");
            sheet.GetValue<string>(1, 3).Should().Be("Header 5", "Header in C1 is incorrect.");
            sheet.GetValue<string>(1, 4).Should().Be("Header 3", "Header in D1 is incorrect.");
            sheet.GetValue<string>(1, 5).Should().Be("Value2", "Header in E1 is incorrect.");
            sheet.GetValue<double>(2, 1).Should().Be(data.Value4, "Value in A2 is incorrect.");
            sheet.Cells[2, 1].Style.Numberformat.Format.Should().Be("???.???", "NumberFormat of A2 is incorrect.");
            sheet.GetValue<string>(2, 2).Should().Be(data.Value1, "Value in B2 is incorrect.");
            sheet.GetValue<string>(2, 3).Should().Be(data.Value5.ToString(), "Value in C2 is incorrect.");
            sheet.GetValue<string>(2, 4).Should().Be(data.Value3.ToString(), "Value in D2 is incorrect.");
            sheet.GetValue<DateTime>(2, 5).Should().Be(data.Value2, "Value in E2 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithListOfComplexTestItem_WritesExcelDocumentToStream()
        {
            var data = new[] { new ComplexTestItem { Value1 = "Item 1",
                                                     Value2 = DateTime.Today,
                                                     Value3 = true,
                                                     Value4 = 100.1,
                                                     Value5 = TestEnum.First,
                                                     Value6 = "Ignored" },

                               new ComplexTestItem { Value1 = "Item 2",
                                                     Value2 = DateTime.Today.AddDays(1),
                                                     Value3 = false,
                                                     Value4 = 200.2,
                                                     Value5 = TestEnum.Second,
                                                     Value6 = "Also ignored" } }.ToList();

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(3, "Worksheet should have three rows (including header column).");
            sheet.Dimension.End.Column.Should().Be(5, "Worksheet should have five columns.");
            sheet.GetValue<string>(1, 1).Should().Be("Header 4", "Header in A1 is incorrect.");
            sheet.GetValue<string>(1, 2).Should().Be("Value1", "Header in B1 is incorrect.");
            sheet.GetValue<string>(1, 3).Should().Be("Header 5", "Header in C1 is incorrect.");
            sheet.GetValue<string>(1, 4).Should().Be("Header 3", "Header in D1 is incorrect.");
            sheet.GetValue<string>(1, 5).Should().Be("Value2", "Header in E1 is incorrect.");
            sheet.GetValue<double>(2, 1).Should().Be(data[0].Value4, "Value in A2 is incorrect.");
            sheet.Cells[2, 1].Style.Numberformat.Format.Should().Be("???.???", "NumberFormat of A2 is incorrect.");
            sheet.GetValue<string>(2, 2).Should().Be(data[0].Value1, "Value in B2 is incorrect.");
            sheet.GetValue<string>(2, 3).Should().Be(data[0].Value5.ToString(), "Value in C2 is incorrect.");
            sheet.GetValue<string>(2, 4).Should().Be(data[0].Value3.ToString(), "Value in D2 is incorrect.");
            sheet.GetValue<DateTime>(2, 5).Should().Be(data[0].Value2, "Value in E2 is incorrect.");
            sheet.GetValue<double>(3, 1).Should().Be(data[1].Value4, "Value in A3 is incorrect.");
            sheet.Cells[3, 1].Style.Numberformat.Format.Should().Be("???.???", "NumberFormat of A3 is incorrect.");
            sheet.GetValue<string>(3, 2).Should().Be(data[1].Value1, "Value in B3 is incorrect.");
            sheet.GetValue<string>(3, 3).Should().Be(data[1].Value5.ToString(), "Value in C3 is incorrect.");
            sheet.GetValue<string>(3, 4).Should().Be(data[1].Value3.ToString(), "Value in D3 is incorrect.");
            sheet.GetValue<DateTime>(3, 5).Should().Be(data[1].Value2, "Value in E3 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithAnonymousObject_WritesExcelDocumentToStream()
        {
            var data = new { prop1 = "val1", prop2 = 1.0, prop3 = DateTime.Today };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(2, "Worksheet should have two rows.");
            sheet.Dimension.End.Column.Should().Be(3, "Worksheet should have three columns.");
            sheet.GetValue<string>(1, 1).Should().Be("prop1", "Header in A1 is incorrect.");
            sheet.GetValue<string>(1, 2).Should().Be("prop2", "Header in B1 is incorrect.");
            sheet.GetValue<string>(1, 3).Should().Be("prop3", "Header in C1 is incorrect.");
            sheet.GetValue<string>(2, 1).Should().Be(data.prop1, "Value in A2 is incorrect.");
            sheet.GetValue<double>(2, 2).Should().Be(data.prop2, "Value in B2 is incorrect.");
            sheet.GetValue<DateTime>(2, 3).Should().Be(data.prop3, "Value in C2 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithArrayOfAnonymousObject_WritesExcelDocumentToStream()
        {
            var data = new[] {
                new { prop1 = "val1", prop2 = 1.0, prop3 = DateTime.Today },
                new { prop1 = "val2", prop2 = 2.0, prop3 = DateTime.Today.AddDays(1) }
            };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(3, "Worksheet should have three rows.");
            sheet.Dimension.End.Column.Should().Be(3, "Worksheet should have three columns.");
            sheet.GetValue<string>(1, 1).Should().Be("prop1", "Header in A1 is incorrect.");
            sheet.GetValue<string>(1, 2).Should().Be("prop2", "Header in B1 is incorrect.");
            sheet.GetValue<string>(1, 3).Should().Be("prop3", "Header in C1 is incorrect.");
            sheet.GetValue<string>(2, 1).Should().Be(data[0].prop1, "Value in A2 is incorrect.");
            sheet.GetValue<double>(2, 2).Should().Be(data[0].prop2, "Value in B2 is incorrect.");
            sheet.GetValue<DateTime>(2, 3).Should().Be(data[0].prop3, "Value in C2 is incorrect.");
            sheet.GetValue<string>(3, 1).Should().Be(data[1].prop1, "Value in A3 is incorrect.");
            sheet.GetValue<double>(3, 2).Should().Be(data[1].prop2, "Value in B3 is incorrect.");
            sheet.GetValue<DateTime>(3, 3).Should().Be(data[1].prop3, "Value in C3 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithString_WritesExcelDocumentToStream()
        {
            var data = "Test";

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(1, "Worksheet should have one row.");
            sheet.Dimension.End.Column.Should().Be(1, "Worksheet should have one column.");
            sheet.GetValue<string>(1, 1).Should().Be(data, "Value in A1 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithArrayOfString_WritesExcelDocumentToStream()
        {
            var data = new[] { "1,1", "2,1" };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(2, "Worksheet should have two rows.");
            sheet.Dimension.End.Column.Should().Be(1, "Worksheet should have one column.");
            sheet.GetValue<string>(1, 1).Should().Be(data[0], "Value in A1 is incorrect.");
            sheet.GetValue<string>(2, 1).Should().Be(data[1], "Value in A2 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithInt32_WritesExcelDocumentToStream()
        {
            var data = 100;

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(1, "Worksheet should have one row.");
            sheet.Dimension.End.Column.Should().Be(1, "Worksheet should have one column.");
            sheet.GetValue<int>(1, 1).Should().Be(data, "Value in A1 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithArrayOfInt32_WritesExcelDocumentToStream()
        {
            var data = new[] { 100, 200 };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(2, "Worksheet should have two rows.");
            sheet.Dimension.End.Column.Should().Be(1, "Worksheet should have one column.");
            sheet.GetValue<int>(1, 1).Should().Be(data[0], "Value in A1 is incorrect.");
            sheet.GetValue<int>(2, 1).Should().Be(data[1], "Value in A2 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithDateTime_WritesExcelDocumentToStream()
        {
            var data = DateTime.Today;

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(1, "Worksheet should have one row.");
            sheet.Dimension.End.Column.Should().Be(1, "Worksheet should have one column.");
            sheet.GetValue<DateTime>(1, 1).Should().Be(data, "Value in A1 is incorrect.");
        }

        [Fact]
        public void WriteToStreamAsync_WithArrayOfDateTime_WritesExcelDocumentToStream()
        {
            var data = new[] { DateTime.Today, DateTime.Today.AddDays(1) };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            sheet.Dimension.Should().NotBeNull("Worksheet has no cells.");
            sheet.Dimension.End.Row.Should().Be(2, "Worksheet should have two rows.");
            sheet.Dimension.End.Column.Should().Be(1, "Worksheet should have one column.");
            sheet.GetValue<DateTime>(1, 1).Should().Be(data[0], "Value in A1 is incorrect.");
            sheet.GetValue<DateTime>(2, 1).Should().Be(data[1], "Value in A2 is incorrect.");
        }

        [Fact]
        public void XlsxMediaTypeFormatter_WithDefaultHeaderHeight_DefaultsToSameHeightForAllCells()
        {
            var data = new[] { new SimpleTestItem { Value1 = "A1", Value2 = "B1" },
                               new SimpleTestItem { Value1 = "A1", Value2 = "B2" }  };

            var formatter = new XlsxMediaTypeFormatter();

            var sheet = GetWorksheetFromStream(formatter, data);

            sheet.Row(1).Height.Should().NotBe(0d, "HeaderHeight should not be zero.");
            sheet.Row(1).Height.Should().Be(sheet.Row(2).Height, "HeaderHeight should be the same as other rows.");
        }

        [Fact]
        public void WriteToStreamAsync_WithCellAndHeaderFormats_WritesFormattedExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var formatter = new XlsxMediaTypeFormatter(
                cellStyle: (ExcelStyle s) =>
                {
                    s.Font.Size = 15f;
                    s.Font.Bold = true;
                },
                headerStyle: (ExcelStyle s) =>
                {
                    s.Font.Size = 18f;
                    s.Border.Bottom.Style = ExcelBorderStyle.Thick;
                }
            );

            var sheet = GetWorksheetFromStream(formatter, data);

            sheet.Cells[1, 1].Style.Font.Bold.Should().BeTrue("Header in A1 should be bold.");
            sheet.Cells[1, 2].Style.Font.Bold.Should().BeTrue("Header in B1 should be bold.");
            sheet.Cells[1, 3].Style.Font.Bold.Should().BeTrue("Header in C1 should be bold.");
            sheet.Cells[2, 1].Style.Font.Bold.Should().BeTrue("Value in A2 should be bold.");
            sheet.Cells[2, 2].Style.Font.Bold.Should().BeTrue("Value in B2 should be bold.");
            sheet.Cells[2, 3].Style.Font.Bold.Should().BeTrue("Value in C2 should be bold.");
            sheet.Cells[3, 1].Style.Font.Bold.Should().BeTrue("Value in A3 should be bold.");
            sheet.Cells[3, 2].Style.Font.Bold.Should().BeTrue("Value in B3 should be bold.");
            sheet.Cells[3, 3].Style.Font.Bold.Should().BeTrue("Value in C3 should be bold.");
            sheet.Cells[1, 1].Style.Font.Size.Should().Be(18f, "Header in A1 should be in size 18 font.");
            sheet.Cells[1, 2].Style.Font.Size.Should().Be(18f, "Header in B1 should be in size 18 font.");
            sheet.Cells[1, 3].Style.Font.Size.Should().Be(18f, "Header in C1 should be in size 18 font.");
            sheet.Cells[2, 1].Style.Font.Size.Should().Be(15f, "Value in A2 should be in size 15 font.");
            sheet.Cells[2, 2].Style.Font.Size.Should().Be(15f, "Value in B2 should be in size 15 font.");
            sheet.Cells[2, 3].Style.Font.Size.Should().Be(15f, "Value in C2 should be in size 15 font.");
            sheet.Cells[3, 1].Style.Font.Size.Should().Be(15f, "Value in A3 should be in size 15 font.");
            sheet.Cells[3, 2].Style.Font.Size.Should().Be(15f, "Value in B3 should be in size 15 font.");
            sheet.Cells[3, 3].Style.Font.Size.Should().Be(15f, "Value in C3 should be in size 15 font.");
            sheet.Cells[1, 1].Style.Border.Bottom.Style.Should().Be(ExcelBorderStyle.Thick, "Header in A1 should have a thick border.");
            sheet.Cells[1, 2].Style.Border.Bottom.Style.Should().Be(ExcelBorderStyle.Thick, "Header in B1 should have a thick border.");
            sheet.Cells[1, 3].Style.Border.Bottom.Style.Should().Be(ExcelBorderStyle.Thick, "Header in C1 should have a thick border.");
            sheet.Cells[2, 1].Style.Border.Bottom.Style.Should().NotBe(ExcelBorderStyle.Thick, "Value in A2 should NOT have a thick border.");
            sheet.Cells[2, 2].Style.Border.Bottom.Style.Should().NotBe(ExcelBorderStyle.Thick, "Value in B2 should NOT have a thick border.");
            sheet.Cells[2, 3].Style.Border.Bottom.Style.Should().NotBe(ExcelBorderStyle.Thick, "Value in C2 should NOT have a thick border.");
            sheet.Cells[3, 1].Style.Border.Bottom.Style.Should().NotBe(ExcelBorderStyle.Thick, "Value in A3 should NOT have a thick border.");
            sheet.Cells[3, 2].Style.Border.Bottom.Style.Should().NotBe(ExcelBorderStyle.Thick, "Value in B3 should NOT have a thick border.");
            sheet.Cells[3, 3].Style.Border.Bottom.Style.Should().NotBe(ExcelBorderStyle.Thick, "Value in C3 should NOT have a thick border.");
        }

        [Fact]
        public void WriteToStreamAsync_WithHeaderRowHeight_WritesFormattedExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var formatter = new XlsxMediaTypeFormatter(headerHeight: 30f);

            var sheet = GetWorksheetFromStream(formatter, data);

            sheet.Row(1).Height.Should().Be(30f, "Row 1 should have height 30.");
        }

        private static ExcelWorksheet GetWorksheetFromStream<TItem>(XlsxMediaTypeFormatter formatter, TItem data)
        {

            var content = new FakeContent();
            content.Headers.ContentType = new MediaTypeHeaderValue("application/atom+xml");

            var context = new OutputFormatterWriteContext(new DefaultHttpContext(), new TestHttpResponseStreamWriterFactory().CreateWriter, typeof(TItem), data);
            formatter.WriteResponseBodyAsync(context, Encoding.UTF8).GetAwaiter().GetResult();

            using (var ms = new MemoryStream())
            {
                ms.Seek(0, SeekOrigin.Begin);
                using (var package = new ExcelPackage(ms))
                {
                    return package.Workbook.Worksheets[1];
                }
            }
        }
    }

    public class TestHttpResponseStreamWriterFactory : IHttpResponseStreamWriterFactory
    {
        public const int DefaultBufferSize = 16 * 1024;

        public TextWriter CreateWriter(Stream stream, Encoding encoding)
        {
            return new HttpResponseStreamWriter(stream, encoding, DefaultBufferSize);
        }
    }
}
