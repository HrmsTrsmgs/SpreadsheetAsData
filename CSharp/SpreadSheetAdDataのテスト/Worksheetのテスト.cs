﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Marimo.SpreadSheetAsData;
using NUnit.Framework;
using NUnit.Framework.Constraints;

namespace Marimo.SpreadSheetAdData.Test
{
    [TestFixture]
    public class Worksheetのテスト : AssertionHelper
    {

        Workbook book;
        Worksheet sheet1;
        Worksheet data;

        [SetUp]
        public void SetUp()
        {
            book = Workbook.Open(@"TestData\Book1.xlsx");
            sheet1 = book.Sheets["Sheet1"];
            data = book.Sheets["いろいろなデータ"];
        }

        [TearDown]
        public void TearDown()
        {
            book.Close();
        }

        [Test]
        public void Nameはシート名を返します()
        {
            Expect(sheet1.Name, Is.EqualTo("Sheet1"));
        }

        [Test]
        public void Nameは日本語でしていたシート名も適切に扱います()
        {
            Expect(data.Name, Is.EqualTo("いろいろなデータ"));
        }

        [Test]
        public void Bookは所属しているWorkbookを取得します()
        {
            Expect(sheet1.Book, Is.SameAs(book));
        }

        [Test]
        public void Cellsはセルの参照文字列を文字列を指定してセル取得します()
        {
            Expect(sheet1.Cells["C3"].Reference, Is.EqualTo("C3"));
        }

        [Test]
        public void Cellsはセルの座標を数値で指定してセル取得します()
        {
            Expect(sheet1.Cells[3, 3].Reference, Is.EqualTo("C3"));
        }
    }
}
