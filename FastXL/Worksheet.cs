using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;

namespace AppAsToy.FastXL
{
	[DebuggerDisplay("Name={Name}, Rows={RowCount}, Columns={ColumnCount}, Cells={CellCount}")]
	public sealed class Worksheet
	{
		readonly WorkbookContext context;

		Row[] rows;

		public readonly string Name;
		public readonly int Index;

		public bool IsLoaded => rows != null;
		public int RowCount
		{
			get
			{
				ThrowIfNotLoaded();
				return rows.Length;
			}
		}

		public int ColumnCount
		{
			get
			{
				ThrowIfNotLoaded();
				return rows.Length > 0 ? rows[0].Count : 0;
			}
		}
		public int CellCount => RowCount * ColumnCount;

		public IReadOnlyList<Row> Rows => rows;


		internal Worksheet(int index, string name, WorkbookContext context)
		{
			Index = index;
			Name = name;
			this.context = context;
		}

		public object Cell(int row, int column)
		{
			ThrowIfNotLoaded();
			return rows[row][column].Value;
		}

		public string ReadXml()
		{
			return context.Archive.ReadString($"sheet{Index + 1}.xml");
		}

		public void Load()
		{
			if (IsLoaded)
				return;

			var sheetXml = ReadXml();
			LoadInternal(sheetXml);
		}

		public async Task LoadAsync()
		{
			if (IsLoaded)
				return;

			var sheetXml = ReadXml();
			await LoadInternalAsync(sheetXml);
		}

		internal void LoadInternal(string sheetXml)
		{
			rows = WorksheetParser.ParseRows(sheetXml, context);
		}

		internal async Task LoadInternalAsync(string sheetXml)
		{
			rows = await Task.Run(() => WorksheetParser.ParseRows(sheetXml, context));
		}

		void ThrowIfNotLoaded()
		{
			if (!IsLoaded)
				throw new InvalidOperationException($"Worksheet({Name}) is not loaded");
		}
	}
}
