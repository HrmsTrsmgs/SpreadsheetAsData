using Marimo.SpreadSheetAsData;
using NUnit.Framework;
using System;
using System.Collections.Generic;
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
        public void 比較するとintの0と同じとされます()
        {
            Assert.That(tested == 0, Is.True);
        }

        [Test]
        public void 比較するとdoubleの0と同じとされます()
        {
            Assert.That(tested == .0, Is.True);
        }

        [Test]
        public void 比較すると空白以外の文字列とは同じとされません()
        {
            Assert.That(tested == "a", Is.False);
        }

        [Test]
        public void 比較すると0以外のintとは同じとされません()
        {
            Assert.That(tested == 1, Is.False);
        }

        [Test]
        public void 比較すると0以外のdoubleとは同じとされません()
        {
            Assert.That(tested == 0.01, Is.False);
        }
    }
}
