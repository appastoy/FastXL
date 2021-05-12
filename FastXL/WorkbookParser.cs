using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FastXL
{
	static class WorkbookParser
	{
		static readonly ParallelOptions parallelOption = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };
		static readonly Regex sheetNameRE = new Regex(@"<sheet name=""([^""]+)"".+?r:id=""rId(\d+)""", RegexOptions.Compiled);
		static readonly Regex sharedStringRE = new Regex(@"<si>(.+?)</si>", RegexOptions.Compiled | RegexOptions.Singleline);
		static readonly Regex sharedStringValuesRE = new Regex(@"<t.*?>(.*?)</t>", RegexOptions.Compiled | RegexOptions.Singleline);
		static readonly Regex stylesRE = new Regex(@"<cellXfs count=""\d+"">(.*?)</cellXfs>", RegexOptions.Compiled | RegexOptions.Singleline);
		static readonly Regex styleItemRE = new Regex(@"<xf numFmtId=""(\d+)""", RegexOptions.Compiled | RegexOptions.Singleline);

		public static async Task<WorkbookContext> ParseAsync(Zip zip)
		{
			var sharedStringsXml = zip.ReadString("sharedStrings.xml");
			var sharedStringsTask = ParseSharedStringsAsync(sharedStringsXml);
			
			var stylesXml = zip.ReadString("styles.xml");
			var stylesTask = ParseStylesAsync(stylesXml);
			
			var workbookXml = zip.ReadString("workbook.xml");
			var sheetNames = ParseSheetNames(workbookXml);

			var sharedStrings = await sharedStringsTask;
			var styles = await stylesTask;

			return new WorkbookContext(zip, sheetNames, sharedStrings, styles);
		}

		static string[] ParseSheetNames(string workbookXml)
		{
			var matches = sheetNameRE.Matches(workbookXml);
			var sheets = new string[matches.Count];
			foreach (Match match in matches)
			{
				var sheetName = match.Groups[1].Value;
				var sheetIndex = int.Parse(match.Groups[2].Value) - 1;
				sheets[sheetIndex] = sheetName;
			}
			return sheets;
		}

		static async Task<string[]> ParseSharedStringsAsync(string sharedStringsXml)
		{
			if (string.IsNullOrEmpty(sharedStringsXml))
				return Array.Empty<string>();

			var matches = sharedStringRE.Matches(sharedStringsXml);
			var sharedStrings = new string[matches.Count];
			await Task.Run(() =>
			{
				Parallel.For(0, sharedStrings.Length, parallelOption, i =>
				{
					var valueMatches = sharedStringValuesRE.Matches(matches[i].Groups[1].Value);
					if (valueMatches.Count == 0)
					{
						sharedStrings[i] = string.Empty;
					}
					else
					{
						sharedStrings[i] = string.Join(string.Empty, 
							valueMatches.OfType<Match>().Select(m => m.Groups[1].Value));
					}
				});
			});
			return sharedStrings;
		}

		static async Task<Style[]> ParseStylesAsync(string stylesXml)
		{
			if (string.IsNullOrEmpty(stylesXml))
				return Array.Empty<Style>();

			var stylesMatch = stylesRE.Match(stylesXml);
			var styleItemMatches = styleItemRE.Matches(stylesMatch.Groups[1].Value);
			var styles = new Style[styleItemMatches.Count];
			
			await Task.Run(() =>
			{
				Parallel.For(0, styles.Length, parallelOption, i =>
				{
					var numberFormat = int.Parse(styleItemMatches[i].Groups[1].Value);
					styles[i] = new Style(numberFormat);
				});
			});
			return styles;
		}
	}
}
