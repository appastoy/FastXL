using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FastXL
{
	public static class ExcelFile
	{
		static readonly Regex sheetNameRE = new Regex(@"<sheet name=""([^""]+)"" sheetId=""(\d+)""", RegexOptions.Compiled);
		static readonly Regex sharedStringRE = new Regex(@"<t(?:>| [^>]+>)([^<]+)</t>", RegexOptions.Compiled | RegexOptions.Multiline);

		public static async Task<Workbook> LoadBookAsync(string xlPath, bool loadAllSheet = false)
		{
			var zip = await Zip.LoadAsync(xlPath);
			var sheetNames = ParseSheetNames(zip);
			var sharedStrings = ParseSharedStrings(zip);

			var workbook = new Workbook(new ExcelContext(zip, sheetNames, sharedStrings));
			if (loadAllSheet)
				await workbook.LoadAllSheetAsync();
			return workbook;
		}

		public static Workbook LoadBook(string xlPath, bool loadAllSheet = false)
		{
			return LoadBookAsync(xlPath, loadAllSheet).Result;
		}

		static string[] ParseSheetNames(Zip zip)
		{
			var workbookXml = zip.Entries.FirstOrDefault(e => e.Name == "workbook.xml")?.ReadAllText();
			if (workbookXml == null)
				throw new InvalidOperationException("Can't find workbook.xml");

			var matches = sheetNameRE.Matches(workbookXml);
			var sheets = new string[matches.Count];
			foreach (Match match in matches)
			{
				var index = int.Parse(match.Groups[2].Value) - 1;
				sheets[index] = match.Groups[1].Value;
			}

			return sheets;
		}

		static string[] ParseSharedStrings(Zip zip)
		{
			var sharedStringsXml = zip.Entries.FirstOrDefault(e => e.Name == "sharedStrings.xml")?.ReadAllText();
			if (sharedStringsXml == null)
				throw new InvalidOperationException("Can't find sharedStrings.xml");

			var matches = sharedStringRE.Matches(sharedStringsXml);
			return matches.OfType<Match>().Select(m => m.Groups[1].Value).ToArray();
		}

		internal static async Task<object[,]> ParseCellsAsync(int sheetIndex, ExcelContext context)
		{
			var sheetNameWithIndex = $"sheet{sheetIndex + 1}.xml";
			var entry = context.Archive.Entries.FirstOrDefault(e => e.Name == sheetNameWithIndex);
			if (entry == null)
				throw new InvalidOperationException($"Can't find {sheetNameWithIndex}");

			var xml = await entry.ReadAllTextAsync();
			return CellGrid.Parse(xml, context.SharedStrings);
		}
	}
}
