using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace FastXL
{
	[DebuggerDisplay("Sheets={Sheets.Count}")]
	public sealed class Workbook
	{
		readonly ExcelContext context;
		readonly Worksheet[] worksheets;

		public Worksheet this[int index] => worksheets[index];
		public Worksheet this[string name] => worksheets.FirstOrDefault(ws => ws.Name == name);
		public IReadOnlyList<Worksheet> Sheets => worksheets;

		internal Workbook(ExcelContext context)
		{
			this.context = context;
			worksheets = context.SheetNames.Select((name, index) => new Worksheet(index, name, this.context)).ToArray();
		}

		public async Task LoadAllSheetAsync()
		{
			var loadSheetTasks = worksheets.Select(ws => ws.LoadAsync()).ToArray();
			await Task.WhenAll(loadSheetTasks);
		}

		public void LoadAllSheet()
		{
			LoadAllSheetAsync().Wait();
		}
	}
}
