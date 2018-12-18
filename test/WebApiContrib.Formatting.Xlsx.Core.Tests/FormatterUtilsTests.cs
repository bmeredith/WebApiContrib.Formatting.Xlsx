using System;
using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using WebApiContrib.Formatting.Xlsx.Core.Tests.TestData;
using Xunit;

namespace WebApiContrib.Formatting.Xlsx.Core.Tests
{
    public class FormatterUtilsTests
    {
        [Fact]
        public void GetAttribute_ExcelColumnAttributeOfComplexTestItemValue2_ExcelColumnAttribute()
        {
            var value2 = typeof(ComplexTestItem).GetMember("Value2")[0];
            var excelAttribute = FormatterUtils.GetAttribute<ExcelColumnAttribute>(value2);

            excelAttribute.Should().NotBeNull();
            excelAttribute.Order.Should().Be(2);
        }

        [Fact]
        public void GetAttribute_ExcelDocumentAttributeOfComplexTestItem_ExcelDocumentAttribute()
        {
            var complexTestItem = typeof(ComplexTestItem);
            var excelAttribute = FormatterUtils.GetAttribute<ExcelDocumentAttribute>(complexTestItem);

            excelAttribute.Should().NotBeNull();
            excelAttribute.FileName.Should().Be("Complex test item");
        }

        [Fact]
        public void MemberOrder_SimpleTestItem_ReturnsMemberOrder()
        {
            var testItemType = typeof(SimpleTestItem);
            var value1 = testItemType.GetMember("Value1")[0];
            var value2 = testItemType.GetMember("Value2")[0];

            FormatterUtils.MemberOrder(value1).Should().Be(-1, "Value1 should have order -1.");
            FormatterUtils.MemberOrder(value2).Should().Be(-1, "Value2 should have order -1.");
        }

        [Fact]
        public void MemberOrder_ComplexTestItem_ReturnsMemberOrder()
        {
            var testItemType = typeof(ComplexTestItem);
            var value1 = testItemType.GetMember("Value1")[0];
            var value2 = testItemType.GetMember("Value2")[0];
            var value3 = testItemType.GetMember("Value3")[0];
            var value4 = testItemType.GetMember("Value4")[0];
            var value5 = testItemType.GetMember("Value5")[0];
            var value6 = testItemType.GetMember("Value6")[0];

            FormatterUtils.MemberOrder(value1).Should().Be(-1, "Value1 should have order -1.");
            FormatterUtils.MemberOrder(value2).Should().Be(2, "Value2 should have order 2.");
            FormatterUtils.MemberOrder(value3).Should().Be(1, "Value3 should have order 1.");
            FormatterUtils.MemberOrder(value4).Should().Be(-2, "Value4 should have order -2.");
            FormatterUtils.MemberOrder(value5).Should().Be(-1, "Value5 should have order -1.");
            FormatterUtils.MemberOrder(value6).Should().Be(-1, "Value6 should have order -1.");
        }

        [Fact]
        public void GetMemberNames_SimpleTestItem_ReturnsMemberNamesInOrder()
        {
            var memberNames = FormatterUtils.GetMemberNames(typeof(SimpleTestItem));

            memberNames.Should().NotBeNull();
            memberNames.Count.Should().Be(2);
            memberNames[0].Should().Be("Value1");
            memberNames[1].Should().Be("Value2");
        }

        [Fact]
        public void GetMemberNames_ComplexTestItem_ReturnsMemberNamesInOrder()
        {
            var memberNames = FormatterUtils.GetMemberNames(typeof(ComplexTestItem));

            memberNames.Should().NotBeNull();
            memberNames.Count.Should().Be(5);
            memberNames[0].Should().Be("Value4");
            memberNames[1].Should().Be("Value1");
            memberNames[2].Should().Be("Value5");
            memberNames[3].Should().Be("Value3");
            memberNames[4].Should().Be("Value2");
        }

        [Fact]
        public void GetMemberNames_AnonymousType_ReturnsMemberNamesInOrderDefined()
        {
            var anonymous = new { prop1 = "value1", prop2 = "value2" };
            var memberNames = FormatterUtils.GetMemberNames(anonymous.GetType());

            memberNames.Should().NotBeNull();
            memberNames.Count.Should().Be(2);
            memberNames[0].Should().Be("prop1");
            memberNames[1].Should().Be("prop2");
        }

        [Fact]
        public void GetMemberInfo_SimpleTestItem_ReturnsMemberInfoList()
        {
            var memberInfo = FormatterUtils.GetMemberInfo(typeof(SimpleTestItem));

            memberInfo.Should().NotBeNull();
            memberInfo.Count.Should().Be(2);
        }

        [Fact]
        public void GetMemberInfo_AnonymousType_ReturnsMemberInfoList()
        {
            var anonymous = new { prop1 = "value1", prop2 = "value2" };
            var memberInfo = FormatterUtils.GetMemberInfo(anonymous.GetType());

            memberInfo.Should().NotBeNull();
            memberInfo.Count.Should().Be(2);
        }

        [Fact]
        public void GetEnumerableItemType_ListOfSimpleTestItem_ReturnsTestItemType()
        {
            var testItemList = typeof(List<SimpleTestItem>);
            var itemType = FormatterUtils.GetEnumerableItemType(testItemList);

            itemType.Should().NotBeNull();
            itemType.Should().Be(typeof(SimpleTestItem));
        }

        [Fact]
        public void GetEnumerableItemType_IEnumerableOfSimpleTestItem_ReturnsTestItemType()
        {
            var testItemList = typeof(IEnumerable<SimpleTestItem>);
            var itemType = FormatterUtils.GetEnumerableItemType(testItemList);

            itemType.Should().NotBeNull();
            itemType.Should().Be(typeof(SimpleTestItem));
        }

        [Fact]
        public void GetEnumerableItemType_ArrayOfSimpleTestItem_ReturnsTestItemType()
        {
            var testItemArray = typeof(SimpleTestItem[]);
            var itemType = FormatterUtils.GetEnumerableItemType(testItemArray);

            itemType.Should().NotBeNull();
            itemType.Should().Be(typeof(SimpleTestItem));
        }

        [Fact]
        public void GetEnumerableItemType_ArrayOfAnonymousObject_ReturnsTestItemType()
        {
            var anonymous = new { prop1 = "value1", prop2 = "value2" };
            var anonymousArray = new[] { anonymous };

            var itemType = FormatterUtils.GetEnumerableItemType(anonymousArray.GetType());

            itemType.Should().NotBeNull();
            itemType.Should().Be(anonymous.GetType());
        }

        [Fact]
        public void GetEnumerableItemType_ListOfAnonymousObject_ReturnsTestItemType()
        {
            var anonymous = new { prop1 = "value1", prop2 = "value2" };
            var anonymousList = new[] { anonymous }.ToList();

            var itemType = FormatterUtils.GetEnumerableItemType(anonymousList.GetType());

            itemType.Should().NotBeNull();
            itemType.Should().Be(anonymous.GetType());
        }

        [Fact]
        public void GetFieldOrPropertyValue_ComplexTestItem_ReturnsPropertyValues()
        {
            var obj = new ComplexTestItem
            {
                Value1 = "Value 1",
                Value2 = DateTime.Today,
                Value3 = true,
                Value4 = 100.1,
                Value5 = TestEnum.Second,
                Value6 = "Value 6"
            };

            FormatterUtils.GetFieldOrPropertyValue(obj, "Value1").Should().Be(obj.Value1);
            FormatterUtils.GetFieldOrPropertyValue(obj, "Value2").Should().Be(obj.Value2);
            FormatterUtils.GetFieldOrPropertyValue(obj, "Value3").Should().Be(obj.Value3);
            FormatterUtils.GetFieldOrPropertyValue(obj, "Value4").Should().Be(obj.Value4);
            FormatterUtils.GetFieldOrPropertyValue(obj, "Value5").Should().Be(obj.Value5);
            FormatterUtils.GetFieldOrPropertyValue(obj, "Value6").Should().Be(obj.Value6);
        }

        [Fact]
        public void GetFieldOrPropertyValueT_ComplexTestItem_ReturnsPropertyValues()
        {
            var obj = new ComplexTestItem
            {
                Value1 = "Value 1",
                Value2 = DateTime.Today,
                Value3 = true,
                Value4 = 100.1,
                Value5 = TestEnum.Second,
                Value6 = "Value 6"
            };

            FormatterUtils.GetFieldOrPropertyValue<string>(obj, "Value1").Should().Be(obj.Value1);
            FormatterUtils.GetFieldOrPropertyValue<DateTime>(obj, "Value2").Should().Be(obj.Value2);
            FormatterUtils.GetFieldOrPropertyValue<bool>(obj, "Value3").Should().Be(obj.Value3);
            FormatterUtils.GetFieldOrPropertyValue<double>(obj, "Value4").Should().Be(obj.Value4);
            FormatterUtils.GetFieldOrPropertyValue<TestEnum>(obj, "Value5").Should().Be(obj.Value5);
            FormatterUtils.GetFieldOrPropertyValue<string>(obj, "Value6").Should().Be(obj.Value6);
        }

        [Fact]
        public void GetFieldOrPropertyValue_AnonymousObject_ReturnsPropertyValues()
        {
            var obj = new { prop1 = "test", prop2 = 2.0, prop3 = DateTime.Today };

            FormatterUtils.GetFieldOrPropertyValue(obj, "prop1").Should().Be(obj.prop1);
            FormatterUtils.GetFieldOrPropertyValue(obj, "prop2").Should().Be(obj.prop2);
            FormatterUtils.GetFieldOrPropertyValue(obj, "prop3").Should().Be(obj.prop3);
        }

        [Fact]
        public void GetFieldOrPropertyValueT_AnonymousObject_ReturnsPropertyValues()
        {
            var obj = new { prop1 = "test", prop2 = 2.0, prop3 = DateTime.Today };

            FormatterUtils.GetFieldOrPropertyValue<string>(obj, "prop1").Should().Be(obj.prop1);
            FormatterUtils.GetFieldOrPropertyValue<double>(obj, "prop2").Should().Be(obj.prop2);
            FormatterUtils.GetFieldOrPropertyValue<DateTime>(obj, "prop3").Should().Be(obj.prop3);
        }

        [Fact]
        public void IsSimpleType_SimpleTypes_ReturnsTrue()
        {
            FormatterUtils.IsSimpleType(typeof(bool)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(byte)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(sbyte)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(char)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(DateTime)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(DateTimeOffset)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(decimal)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(double)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(float)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(Guid)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(int)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(uint)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(long)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(ulong)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(short)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(TimeSpan)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(ushort)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(string)).Should().BeTrue();
            FormatterUtils.IsSimpleType(typeof(TestEnum)).Should().BeTrue();
        }

        [Fact]
        public void IsSimpleType_ComplexTypes_ReturnsFalse()
        {
            var anonymous = new { prop = "val" };

            FormatterUtils.IsSimpleType(anonymous.GetType()).Should().BeFalse();
            FormatterUtils.IsSimpleType(typeof(Array)).Should().BeFalse();
            FormatterUtils.IsSimpleType(typeof(IEnumerable<>)).Should().BeFalse();
            FormatterUtils.IsSimpleType(typeof(object)).Should().BeFalse();
            FormatterUtils.IsSimpleType(typeof(SimpleTestItem)).Should().BeFalse();
        }
    }
}
