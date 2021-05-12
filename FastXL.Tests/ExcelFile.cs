using NUnit.Framework;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace FastXL.Tests
{
	public class ExcelFileTests
	{
		[Test]
		public async Task Empty_Book()
		{
			var book = await Excel.LoadBookAsync("empty.xlsx");
			Assert.That(book, Is.Not.Null);
		}

		[Test]
		public async Task Book()
		{
			var book = await Excel.LoadBookAsync("test.xlsx", false);
			Assert.That(book, Is.Not.Null);
			Assert.That(book.Sheets.All(s => s.IsLoaded == false), Is.True);
		}

		[Test]
		public async Task Book_with_Load()
		{
			var book = await Excel.LoadBookAsync("test.xlsx");
			Assert.That(book, Is.Not.Null);
			Assert.That(book.Sheets.All(s => s.IsLoaded == true), Is.True);
		}

		[Test]
		public async Task Sheet()
		{
			var book = await Excel.LoadBookAsync("test.xlsx");
			foreach (var sheet in book.Sheets)
			{
				Assert.That(sheet, Is.EqualTo(book[sheet.Name]));
				Assert.That(sheet, Is.EqualTo(book[sheet.Index]));
			}
		}

		[Test]
		public async Task Range()
		{
			var book = await Excel.LoadBookAsync("test.xlsx", false);
			var range = book["Range"];
			range.Load();
			Assert.That(range.ColumnCount, Is.EqualTo(9));
			Assert.That(range.RowCount, Is.EqualTo(23));
		}

		[Test]
		public async Task Values()
		{
			var book = await Excel.LoadBookAsync("test.xlsx", false);
			var values = book["Values"];
			values.Load();
			// 1	2.3		abc		TRUE	1%	2012-08-11	2.3		3.3		abcTRUE
			Assert.That(values.Cell(0, 0), Is.EqualTo(1));
			Assert.That(values.Cell(0, 1), Is.EqualTo(2.3));
			Assert.That(values.Cell(0, 2), Is.EqualTo("abc"));
			Assert.That(values.Cell(0, 3), Is.EqualTo(true)); // true
			Assert.That(values.Cell(0, 4), Is.EqualTo(false)); // false
			Assert.That(values.Cell(0, 5), Is.EqualTo(0.01));
			Assert.That(values.Cell(0, 6), Is.EqualTo(new DateTime(2012, 8, 11)));
			Assert.That(values.Cell(0, 7), Is.EqualTo(new DateTime(2021, 5, 12)));
			Assert.That(values.Cell(0, 8), Is.EqualTo(new DateTime(2011, 5, 23, 19, 12, 30)));
			Assert.That(values.Cell(0, 9), Is.EqualTo(2.3));
			Assert.That(values.Cell(0, 10), Is.EqualTo(3.3));
			Assert.That(values.Cell(0, 11), Is.EqualTo("abcTRUE"));
		}
	}
}