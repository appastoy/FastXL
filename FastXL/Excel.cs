using System.IO;
using System.Threading.Tasks;

namespace FastXL
{
	public static class Excel
	{
		public static async Task<Workbook> LoadBookAsync(string xlPath, bool loadAllSheet = true)
		{
			var stream = new FileStream(xlPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, true);
			return await LoadBookAsync(stream, loadAllSheet);
		}

		public static async Task<Workbook> LoadBookAsync(byte[] bytes, bool loadAllSheet = true)
		{
			var stream = new MemoryStream(bytes, false);
			return await LoadBookAsync(stream, loadAllSheet);
		}

		public static async Task<Workbook> LoadBookAsync(Stream stream, bool loadAllSheet = true)
		{
			var zip = Zip.Open(stream);
			var context = await WorkbookParser.ParseAsync(zip);
			var workbook = new Workbook(context);
			if (loadAllSheet)
				await workbook.LoadAllSheetsAsync();
			return workbook;
		}

		public static Workbook LoadBook(string xlPath, bool loadAllSheet = true)
		{
			return LoadBookAsync(xlPath, loadAllSheet).Result;
		}

		public static Workbook LoadBook(byte[] bytes, bool loadAllSheet = true)
		{
			return LoadBookAsync(bytes, loadAllSheet).Result;
		}

		public static Workbook LoadBook(Stream stream, bool loadAllSheet = true)
		{
			return LoadBookAsync(stream, loadAllSheet).Result;
		}
	}
}
