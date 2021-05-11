using System;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace FastXL
{
	[DebuggerDisplay("Row={RowCount}, Columns={ColumnCount}")]
	internal class CellGrid
	{
		const int alphabetRange = 'Z' - 'A' + 1;
		static readonly Regex dimensionRE = new(@"<dimension ref=""([^""]+)""", RegexOptions.Compiled);
		static readonly Regex addressRE = new(@"([A-Z]+)(\d*)", RegexOptions.Compiled);
		static readonly Regex cellRE = new(@"<c r=""([^""]+)""(.*?)>(.+?)</c>", RegexOptions.Compiled);
		static readonly Regex valueRE = new(@"<v>(.+)</v>", RegexOptions.Compiled);
		static readonly Regex floatRE = new(@"^[-+]?\d+(?:\.?[eE][+-]?|\.)\d+$", RegexOptions.Compiled | RegexOptions.Multiline);
		static readonly Regex intRE = new(@"^-?\d+(?:[eE]+?\d+)?$", RegexOptions.Compiled | RegexOptions.Multiline);

		public static object[,] Parse(string xml, string[] sharedStrings)
		{
			var (maxRow, maxColumn) = ParseDimension(xml);
			var cells = new object[maxRow, maxColumn];

			return ParseCells(xml, cells, sharedStrings);
		}

		static (int MaxRow, int MaxColumn) ParseDimension(string xml)
		{
			var match = dimensionRE.Match(xml);
			var last = match.Groups[1].Value.Split(':').Last();
			return ParseAddress(last);
		}

		static (int Row, int Column) ParseAddress(string address)
		{
			var match = addressRE.Match(address);
			var column = ConvertAlphabetToIndex(match.Groups[1].Value);
			var row = int.Parse(match.Groups[2].Value);
			return (row, column);
		}

		static object[,] ParseCells(string xml, object[,] cellGrid, string[] sharedStrings)
		{
			int maxRow = 0;
			int maxCol = 0;
			foreach (Match match in cellRE.Matches(xml))
			{
				var (row, col) = ParseAddress(match.Groups[1].Value);
				if (row > maxRow) maxRow = row;
				if (col > maxCol) maxCol = col;
				cellGrid[row - 1, col - 1] = ParseValue(match.Groups[2].Value, match.Groups[3].Value, sharedStrings);
			}

			if (cellGrid.GetLength(1) != maxRow ||
				cellGrid.GetLength(0) != maxCol)
			{
				var fitCellGrid = new object[maxRow, maxCol];
				for (var r = 0; r < maxRow; ++r)
					for (var c = 0; c < maxCol; ++c)
						fitCellGrid[r, c] = cellGrid[r, c];
				cellGrid = fitCellGrid;
			}

			return cellGrid;
		}

		static object ParseValue(string attrXml, string valueXml, string[] sharedStrings)
		{
			var match = valueRE.Match(valueXml);
			if (!match.Success)
				return valueXml;

			var valueString = match.Groups[1].Value;
			if (attrXml.Contains(@"t=""s"""))
				return sharedStrings[int.Parse(valueString)];

			if (attrXml.Contains(@"t=""b"""))
				return valueString != "0";

			if (attrXml.Contains(@"s=""3"""))
				return DateTime.FromOADate(double.Parse(valueString));

			if (floatRE.IsMatch(valueString))
			{
				return double.TryParse(valueString, out var doubleValue) ? doubleValue :
					   decimal.TryParse(valueString, out var decimalValue) ? decimalValue :
					   0.0;
			}

			if (intRE.IsMatch(valueString))
			{
				return int.TryParse(valueString, out var intValue) ? intValue :
					   long.TryParse(valueString, out var longValue) ? longValue :
					   ulong.TryParse(valueString, out var ulongValue) ? ulongValue :
					   0;
			}

			return bool.TryParse(valueString, out var boolValue) ? boolValue : valueString;
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
	}
}
