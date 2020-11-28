using Marimo.SpreadSheetAsData;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Numerics;
using System.Text;

namespace SpreadSheetAdDataのテスト
{
    [TestFixture]
    public class BlankValueのテスト
    {
        Workbook book;
        Worksheet sheet1;
        Cell cell;
        dynamic tested;

        [SetUp]
        public void SetUp()
        {
            book = Workbook.Open(@"TestData\Book1.xlsx");
            sheet1 = book.Sheets["Sheet1"];
            cell = sheet1.Cells["B1"];
            tested = cell.Value;
        }

        [TearDown]
        public void TearDown()
        {
            book.Close();
        }

        [Test]
        public void 比較すると空白と同じとされます()
        {
            Assert.That(tested == "", Is.True);
        }

        [Test]
        public void 比較するとdoubleの0と同じとされます()
        {
            Assert.That(tested == .0, Is.True);
        }

        [Test]
        public void 文字列として演算すると空白と同じとされます()
        {
            Assert.That(tested + "A", Is.EqualTo("A"));
        }

        [Test]
        public void 数値として加算すると0と同じとされます()
        {
            Assert.That(tested + 3, Is.EqualTo(3));
        }
        [Test]
        public void 数値として乗算すると0と同じとされます()
        {
            Assert.That(tested * 3, Is.EqualTo(0));
        }

        [Test]
        public void ToStringで中かっこに囲まれたBlankとなります()
        {
            Assert.That(tested.ToString(), Is.EqualTo("{Blank}"));
        }
    }
}
