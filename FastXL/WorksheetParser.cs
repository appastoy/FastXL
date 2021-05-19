using System.IO;
using System.Xml;

namespace FastXL
{
	static class WorksheetParser
	{
		public static Row[] ParseRows(string sheetXml, WorkbookContext context)
		{
			var reader = XmlReader.Create(new StringReader(sheetXml));
			var (maxRow, maxColumn) = ParseDimension(reader);
			return ParseCells(reader, maxRow, maxColumn, context);
		}

		static (int MaxRow, int MaxColumn) ParseDimension(XmlReader reader)
		{
			reader.ReadToFollowing("dimension");
			reader.MoveToAttribute("ref");
			var sheetRange = reader.Value;
			var ranges = sheetRange.Split(':');

			return ParsingHelper.ParseAddress(ranges[ranges.Length-1]);
		}
		static Row[] ParseCells(XmlReader reader, int maxRow, int maxColumn, WorkbookContext context)
		{
			var rows = new Row[maxRow];
			for (int i = 0; i < maxRow; i++)
				rows[i] = new Row(new Cell[maxColumn]);

			var usedMaxRow = 0;
			var usedMaxColumn = 0;

			while (reader.ReadToFollowing("c"))
			{
				var address = reader.GetAttribute("r");
				var type = reader.GetAttribute("t");
				var style = reader.GetAttribute("s");
				string value = null;
				if (reader.ReadToDescendant("v") && !reader.IsEmptyElement)
					value = reader.ReadElementContentAsString();
				var (row, column) = ParsingHelper.ParseAddress(address);
				if (row > usedMaxRow) usedMaxRow = row;
				if (column > usedMaxColumn) usedMaxColumn = column;

				var cellValue = CellParser.Parse(type, style, value, context);
				rows[row - 1][column - 1] = new Cell(cellValue);
			}
			
			if (usedMaxRow != maxRow || usedMaxColumn != maxColumn)
				rows = TrimUnusedRange(rows, usedMaxRow, usedMaxColumn);

			return rows;
		}

		static Row[] TrimUnusedRange(Row[] rows, int usedMaxRow, int usedMaxColumn)
		{
			if (rows.Length != usedMaxRow)
			{
				var newRows = new Row[usedMaxRow];
				for (int i = 0; i < usedMaxRow; ++i)
					newRows[i] = rows[i];
				rows = newRows;
			}

			if (usedMaxRow > 0 && rows[0].Count != usedMaxColumn)
			{
				for (int i = 0; i < usedMaxRow; i++)
				{
					var row = rows[i];
					var newCells = new Cell[usedMaxColumn];
					for (int j = 0; j < usedMaxColumn; j++)
						newCells[j] = row[j];
					rows[i] = new Row(newCells);
				}
			}

			return rows;
		}
	}
}
