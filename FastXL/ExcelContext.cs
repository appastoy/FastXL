using System;
using System.IO.Compression;

namespace FastXL
{
	internal sealed class ExcelContext
	{
		public readonly Zip Archive;
		public readonly string[] SheetNames;
		public readonly string[] SharedStrings;

		public ExcelContext(Zip archive, string[] sheetNames, string[] sharedStrings)
		{
			Archive = archive ?? throw new ArgumentNullException(nameof(archive));
			SheetNames = sheetNames ?? throw new ArgumentNullException(nameof(sheetNames));
			SharedStrings = sharedStrings ?? throw new ArgumentNullException(nameof(sharedStrings));
		}
	}
}
