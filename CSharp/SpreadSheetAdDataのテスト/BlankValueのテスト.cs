using Marimo.SpreadSheetAsData;
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

        [Fact]
        public void 比較すると空白と同じとされます()
        {
            Assert.True((bool)(tested == ""));
        }

        [Fact]
        public void 比較するとdoubleの0と同じとされます()
        {
            Assert.True((bool)(tested == .0));
        }

        [Fact]
        public void 文字列として演算すると空白と同じとされます()
        {
            Assert.Equal("A", tested + "A");
        }

        [Fact]
        public void 数値として加算すると0と同じとされます()
        {
            Assert.Equal(3, tested + 3);
        }
        [Fact]
        public void 数値として乗算すると0と同じとされます()
        {
            Assert.Equal(0, tested * 3);
        }

        [Fact]
        public void ToStringで中かっこに囲まれたBlankとなります()
        {
            Assert.Equal("{Blank}", tested.ToString());
        }
    }
}
