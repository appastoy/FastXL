using NUnit.Framework;
using System.Linq;
using System.Threading.Tasks;

namespace AppAsToy.FastXL.Tests
{
	public class WorkbookTests
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
	}
}
