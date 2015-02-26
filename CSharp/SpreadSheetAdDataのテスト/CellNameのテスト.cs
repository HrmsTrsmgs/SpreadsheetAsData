using Marimo.SpreadSheetAsData;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework.Constraints;

namespace Marimo.SpreadSheetAdData.Test
{
    [TestFixture]
    public class CellNameのテスト : AssertionHelper
    {
        [Test]
        public void ColumnNameプロパティで列名を取得できます()
        {
            Expect(new CellName("A1").ColumnName, Is.EqualTo("A"));
            Expect(new CellName("B1").ColumnName, Is.EqualTo("B"));
        }

        [Test]
        public void ColumnIndexプロパティで列番号を取得できます()
        {
            Expect(new CellName("A1").ColumnIndex, Is.EqualTo(1));
            Expect(new CellName("B1").ColumnIndex, Is.EqualTo(2));
        }

        [Test]
        public void RowIndexプロパティで列番号を取得できます()
        {
            Expect(new CellName("A1").RowIndex, Is.EqualTo(1));
            Expect(new CellName("A2").RowIndex, Is.EqualTo(2));
        }
    }
}
