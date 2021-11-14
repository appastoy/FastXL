using System;

namespace AppAsToy.FastXL
{
	class WorkbookContext
	{
		public readonly Zip Archive;
		public readonly string[] SheetNames;
		public readonly string[] SharedStrings;
		public readonly Style[] Styles;

		public WorkbookContext(Zip archive, string[] sheetNames, string[] sharedStrings, Style[] styles)
		{
			Archive = archive ?? throw new ArgumentNullException(nameof(archive));
			SheetNames = sheetNames ?? throw new ArgumentNullException(nameof(sheetNames));
			SharedStrings = sharedStrings ?? throw new ArgumentNullException(nameof(sharedStrings));
			Styles = styles ?? throw new ArgumentNullException(nameof(styles));
		}
	}
}
