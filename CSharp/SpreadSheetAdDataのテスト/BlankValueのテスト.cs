using Marimo.SpreadSheetAsData;
using FluentAssertions;
using Xunit;
using System;
using System.Collections.Generic;
using System.Numerics;
using System.Text;

namespace Marimo.SpreadSheetAdData.Test
{
    public class BlankValueのテスト : IDisposable
    {
        Workbook book;
        Worksheet sheet1;
        Cell cell;
        dynamic tested;

        public BlankValueのテスト()
        {
            book = Workbook.Open(@"TestData\Book1.xlsx");
            sheet1 = book.Sheets["Sheet1"];
            cell = sheet1.Cells["B1"];
            tested = cell.Value;
        }

        public void Dispose()
        {
            book.Close();
        }

        private BlankValue TestedBlankValue =>
            (tested as object).Should().BeOfType<BlankValue>().Which;

        [Fact]
        public void 比較すると空白と同じとされます()
        {
            bool stringと比較した結果 = TestedBlankValue == "";

            stringと比較した結果.Should().BeTrue();
        }

        [Fact]
        public void 比較するとdoubleの0と同じとされます()
        {
            bool doubleと比較した結果 = TestedBlankValue == .0;

            doubleと比較した結果.Should().BeTrue();
        }

        [Fact]
        public void 文字列として演算すると空白と同じとされます()
        {
            (TestedBlankValue + "A").Should().Be("A");
        }

        [Fact]
        public void 数値として加算すると0と同じとされます()
        {
            (TestedBlankValue + 3).Should().Be(3);
        }
        [Fact]
        public void 数値として乗算すると0と同じとされます()
        {
            (TestedBlankValue * 3).Should().Be(0);
        }

        [Fact]
        public void ToStringで中かっこに囲まれたBlankとなります()
        {
            TestedBlankValue.ToString().Should().Be("{Blank}");
        }
    }
}
