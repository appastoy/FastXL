using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace FastXL
{
	[DebuggerDisplay("Sheets={Sheets.Count}")]
	public sealed class Workbook : IDisposable
	{
		readonly WorkbookContext context;
		readonly Worksheet[] worksheets;
		bool isDisposed;

		public Worksheet this[int index] => worksheets[index];
		public Worksheet this[string name] => worksheets.FirstOrDefault(ws => ws.Name == name);
		public IReadOnlyList<Worksheet> Sheets => worksheets;

		internal Workbook(WorkbookContext context)
		{
			this.context = context;
			worksheets = context.SheetNames.Select((name, index) => new Worksheet(index, name, this.context)).ToArray();
		}

		public async Task LoadAllSheetsAsync()
		{
			var loadSheetTasks = new List<Task>(worksheets.Length);
			foreach (var worksheet in worksheets)
			{
				if (worksheet.IsLoaded)
					continue;

				var xml = worksheet.ReadXml();
				var task = worksheet.LoadInternalAsync(xml);
				loadSheetTasks.Add(task);
			}
			if (loadSheetTasks.Count == 0)
				return;
		
			var tasks = loadSheetTasks.ToArray();
			await Task.WhenAll(tasks);
			Dispose();
		}

		public void LoadAllSheets()
		{
			LoadAllSheetsAsync().Wait();
		}

		~Workbook() => Dispose();

		public void Dispose()
		{
			if (!isDisposed)
			{
				context.Archive.Dispose();
				GC.SuppressFinalize(this);
				isDisposed = true;
			}
		}
	}
}
