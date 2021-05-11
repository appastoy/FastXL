using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;

namespace FastXL
{
	[DebuggerDisplay("Name={Name}, Rows={RowCount}, Columns={ColumnCount}, Cells={CellCount}")]
	public sealed class Worksheet
	{
		readonly ExcelContext context;
		internal object[,] cells;

		public readonly string Name;
		public readonly int Index;

		public bool IsLoaded { get; private set; }
		public int RowCount => cells.GetLength(0);
		public int ColumnCount => cells.GetLength(1);
		public int CellCount => cells.Length;


		internal Worksheet(int index, string name, ExcelContext context)
		{
			Index = index;
			Name = name;
			this.context = context;
		}

		public object Cell(int row, int column)
		{
			if (!IsLoaded)
				throw new System.InvalidOperationException($"{Name} worksheet is not loaded.");
			return cells[row, column];
		}

		public async Task LoadAsync()
		{
			if (IsLoaded)
				return;

			cells = await ExcelFile.ParseCellsAsync(Index, context);
			IsLoaded = true;
		}

		public void Load()
		{
			if (IsLoaded)
				return;

			cells = ExcelFile.ParseCellsAsync(Index, context).Result;
			IsLoaded = true;
		}

		public string ToCSV()
		{
			var builder = new StringBuilder(4096);

			var rowCount = RowCount;
			var colCount = ColumnCount;

			for (int ri = 0; ri < rowCount; ri++)
			{
				for (int ci = 0; ci < colCount; ci++)
				{
					var value = ConvertCsvValue(cells[ri, ci]);
					builder.Append(value);
				}
				if (ri + 1 < rowCount)
					builder.AppendLine();
			}

			return builder.ToString();
		}

		static string ConvertCsvValue(object value)
		{
			if (value == null)
				return string.Empty;

			var valueString = value.ToString();
			if (valueString.Length == 0)
				return string.Empty;

			bool needAroundDoubleQuote = false;
			bool hasDoubleQuote = false;
			foreach (var ch in valueString)
			{
				if (ch == ',' || ch == '\n')
				{
					needAroundDoubleQuote = true;
				}
				else if (ch == '"')
				{
					needAroundDoubleQuote = true;
					hasDoubleQuote = true;
					break;
				}
			}

			if (hasDoubleQuote)
				valueString = valueString.Replace("\"", "\"\"");

			if (needAroundDoubleQuote)
				return $"\"{valueString}\",";

			return $"{valueString},";
		}
	}
}
