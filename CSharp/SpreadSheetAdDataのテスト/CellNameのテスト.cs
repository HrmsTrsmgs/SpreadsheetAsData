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
        public void ParseメソッドでCellNameが生成できます()
        {
            Expect(() => CellName.Parse("A1"), Throws.Nothing);
        }

        [Test]
        public void Parseメソッドに無効なセル名を指定するとFormatExceptionとなります()
        {
            Expect(() => CellName.Parse("A"), Throws.InstanceOf<FormatException>());
            Expect(() => CellName.Parse("1"), Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void ColumnNameプロパティで列名を取得できます()
        {
            Expect(CellName.Parse("A1").ColumnName, Is.EqualTo("A"));
            Expect(CellName.Parse("B1").ColumnName, Is.EqualTo("B"));
        }

        [Test]
        public void ColumnIndexプロパティで列番号を取得できます()
        {
            Expect(CellName.Parse("A1").ColumnIndex, Is.EqualTo(1));
            Expect(CellName.Parse("B1").ColumnIndex, Is.EqualTo(2));
        }

        [Test]
        public void RowIndexプロパティで列番号を取得できます()
        {
            Expect(CellName.Parse("A1").RowIndex, Is.EqualTo(1));
            Expect(CellName.Parse("A2").RowIndex, Is.EqualTo(2));
        }

        [Test]
        public void ToStringメソッドはA1形式で文字列を返します()
        {
            Expect(CellName.Parse("A1").ToString(), Is.EqualTo("A1"));
            Expect(CellName.Parse("B2").ToString(), Is.EqualTo("B2"));
        }
    }
}
