using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAdData.Test
{
    public class WorkSheetCollectionのテスト
    {
        WorksheetCollection tested;

        public WorkSheetCollectionのテスト()
        {
            using(var book = Workbook.Open(@"TestData\Book1.xlsx"))
            {
                tested = book.Sheets;
            }
        }

        [Fact]
        public void インデクサに数字でアクセスできます()
        {
            Assert.Equal("Sheet1", tested[0].Name);
        }

        [Fact]
        public void インデクサにシート名でアクセスできます()
        {
            Assert.Same(tested[0], tested["Sheet1"]);
        }

        [Fact]
        public void Countで個数を取得できます()
        {
            Assert.Equal(3, tested.Count);
        }

        [Fact]
        public void Keysでシート名の一覧が取得できます()
        {
            Assert.Contains("Sheet1", tested.Keys);
            Assert.Contains("Sheet2", tested.Keys);
        }

        [Fact]
        public void Valuesでシートの一覧が取得できます()
        {
            Assert.Contains(tested[0], tested.Values);
            Assert.Contains(tested[1], tested.Values);
        }

        [Fact]
        public void ContainsKeyでシート名の有無が確認できます()
        {
            Assert.True((bool)(tested.ContainsKey("Sheet1")));
            Assert.False((bool)(tested.ContainsKey("")));
        }

        [Fact]
        public void TryGetValueでシート名の有無が確認しつつシートの取得ができます()
        {
            Worksheet sheet;
            Assert.True((bool)(tested.TryGetValue("Sheet1", out sheet)));
            Assert.Same(tested[0], sheet);
            Assert.False((bool)(tested.TryGetValue("", out sheet)));
        }

        [Fact]
        public void foreachでシートが取得できます()
        {
            int i = 0;
            foreach (var item in tested)
            {

                Assert.Equal(tested[i++], item);
            }
        }

        [Fact]
        public void 非ジェネリックのforeachでがシートが取得できます()
        {
            int i = 0;
            foreach (var item in (IEnumerable)tested)
            {
                Assert.Equal(tested[i++], item);
            }
        }

        [Fact]
        public void Dictionaryに対するのforeachでがシートが取得できます()
        {
            int i = 0;
            foreach (var item in (IReadOnlyDictionary<string, Worksheet>)tested)
            {
                Assert.Equal(tested[i].Name, item.Key);
                Assert.Equal(tested[i++], item.Value);
            }
        }
    }
}
