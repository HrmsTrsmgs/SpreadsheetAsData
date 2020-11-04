using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Marimo.SpreadSheetAsData;
using NUnit.Framework;

namespace Marimo.SpreadSheetAdData.Test
{
    [TestFixture]
    public class WorkSheetCollectionのテスト
    {
        WorksheetCollection tested;

        [SetUp]
        public void SetUp()
        {
            using(var book = Workbook.Open(@"TestData\Book1.xlsx"))
            {
                tested = book.Sheets;
            }
        }

        [Test]
        public void インデクサに数字でアクセスできます()
        {
            Assert.That(tested[0].Name, Is.EqualTo("Sheet1"));
        }

        [Test]
        public void インデクサにシート名でアクセスできます()
        {
            Assert.That(tested["Sheet1"], Is.SameAs(tested[0]));
        }

        [Test]
        public void Countで個数を取得できます()
        {
            Assert.That(tested.Count, Is.EqualTo(3));
        }

        [Test]
        public void Keysでシート名の一覧が取得できます()
        {
            Assert.That(tested.Keys, Has.Member("Sheet1"));
            Assert.That(tested.Keys, Has.Member("Sheet2"));
        }

        [Test]
        public void Valuesでシートの一覧が取得できます()
        {
            Assert.That(tested.Values, Has.Member(tested[0]));
            Assert.That(tested.Values, Has.Member(tested[1]));
        }

        [Test]
        public void ContainsKeyでシート名の有無が確認できます()
        {
            Assert.That(tested.ContainsKey("Sheet1"), Is.True);
            Assert.That(tested.ContainsKey(""), Is.False);
        }

        [Test]
        public void TryGetValueでシート名の有無が確認しつつシートの取得ができます()
        {
            Worksheet sheet;
            Assert.That(tested.TryGetValue("Sheet1", out sheet), Is.True);
            Assert.That(sheet, Is.SameAs(tested[0]));
            Assert.That(tested.TryGetValue("", out sheet), Is.False);
        }

        [Test]
        public void foreachでシートが取得できます()
        {
            int i = 0;
            foreach (var item in tested)
            {

                Assert.That(item, Is.EqualTo(tested[i++]));
            }
        }

        [Test]
        public void 非ジェネリックのforeachでがシートが取得できます()
        {
            int i = 0;
            foreach (var item in (IEnumerable)tested)
            {
                Assert.That(item, Is.EqualTo(tested[i++]));
            }
        }

        [Test]
        public void Dictionaryに対するのforeachでがシートが取得できます()
        {
            int i = 0;
            foreach (var item in (IReadOnlyDictionary<string, Worksheet>)tested)
            {
                Assert.That(item.Key, Is.EqualTo(tested[i].Name));
                Assert.That(item.Value, Is.EqualTo(tested[i++]));
            }
        }
    }
}
