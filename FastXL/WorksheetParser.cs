using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FastXL
{
	static class WorksheetParser
	{
		const int alphabetRange = 'Z' - 'A' + 1;
		static readonly Regex dimensionRE = new Regex(@"<dimension ref=""([^""]+)""", RegexOptions.Compiled);
		static readonly Regex addressRE = new Regex(@"([a-zA-Z]+)(\d*)", RegexOptions.Compiled);
		static readonly Regex cellRE = new Regex(@"<c r=""([^""]+)""(.*?)>(.+?)</c>", RegexOptions.Compiled);

		public static async Task<Row[]> ParseRowsAsync(string sheetXml, WorkbookContext context)
		{
			var (maxRow, maxColumn) = ParseDimension(sheetXml);
			return await ParseCellsAsync(sheetXml, maxRow, maxColumn, context);
		}

		static (int MaxRow, int MaxColumn) ParseDimension(string xml)
		{
			var match = dimensionRE.Match(xml);
			var usedRange = match.Groups[1].Value;
			var ranges = usedRange.Split(':');

			return ParseAddress(ranges[ranges.Length-1]);
		}

		static (int Row, int Column) ParseAddress(string address)
		{
			var match = addressRE.Match(address);
			var column = ConvertAlphabetToIndex(match.Groups[1].Value);
			var row = int.Parse(match.Groups[2].Value);
			return (row, column);
		}

		static int ConvertAlphabetToIndex(string address)
		{
			var sum = 0;
			var unit = 1;
			for (int i = address.Length - 1; i >= 0; i--)
			{
				var ch = char.ToUpper(address[i]);
				sum += (ch - 'A' + 1) * unit;
				unit *= alphabetRange;
			}
			return sum;
		}

		static async Task<Row[]> ParseCellsAsync(string xml, int maxRow, int maxColumn, WorkbookContext context)
		{
			var rows = new Row[maxRow];
			for (int i = 0; i < maxRow; i++)
				rows[i] = new Row(new Cell[maxColumn]);

			var cellMatches = cellRE.Matches(xml).OfType<Match>();
			await Task.Run(() =>
			{
				var option = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };
				Parallel.ForEach(cellMatches, option, cellMatch =>
				{
					var cellAddress = cellMatch.Groups[1].Value;
					var cellAttribute = cellMatch.Groups[2].Value;
					var cellValue = cellMatch.Groups[3].Value;

					var (row, col) = ParseAddress(cellAddress);
					var value = CellParser.Parse(cellAttribute, cellValue, context);
					rows[row - 1][col - 1] = new Cell(value);
				});
			});

			return rows;
		}
	}
}
