using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AppAsToy.FastXL
{
	static class WorkbookParser
	{
		public static async Task<WorkbookContext> ParseAsync(Zip zip)
		{
			var sharedStringsXml = zip.ReadString("sharedStrings.xml");
			var sharedStringsTask = Task.Run(() => ParseSharedStrings(sharedStringsXml));
			
			var stylesXml = zip.ReadString("styles.xml");
			var stylesTask = Task.Run(() => ParseStyles(stylesXml));
			
			var workbookXml = zip.ReadString("workbook.xml");
			var sheetNames = ParseSheetNames(workbookXml);

			var sharedStrings = await sharedStringsTask;
			var styles = await stylesTask;

			return new WorkbookContext(zip, sheetNames, sharedStrings, styles);
		}

		static string[] ParseSheetNames(string workbookXml)
		{
			var sheets = new List<(int Index, string Name)>(32);
			var reader = XmlReader.Create(new StringReader(workbookXml));
			while (reader.ReadToFollowing("sheet"))
			{
				if (reader.MoveToAttribute("name"))
				{
					var name = reader.Value;
					while (reader.MoveToNextAttribute())
					{
						if (reader.LocalName == "id")
						{
							var index = int.Parse(reader.Value.Substring(3));
							sheets.Add((index, name));
							break;
						}
					}
				}
			}

			return sheets.OrderBy(s => s.Index).Select(s => s.Name).ToArray();
		}

		static string[] ParseSharedStrings(string sharedStringsXml)
		{
			if (string.IsNullOrEmpty(sharedStringsXml))
				return Array.Empty<string>();

			var builder = new StringBuilder(4096);
			var sharedStrings = new List<string>(4096);
			var reader = XmlReader.Create(new StringReader(sharedStringsXml));
			while (reader.ReadToFollowing("si"))
			{
				builder.Length = 0;
				var subReader = reader.ReadSubtree();
				while (subReader.ReadToFollowing("t"))
				{
					if (subReader.IsEmptyElement)
						continue;
					var value = subReader.ReadElementContentAsString() ?? string.Empty;
					builder.Append(value);
				}
				sharedStrings.Add(builder.ToString());
			}

			return sharedStrings.ToArray();
		}

		static Style[] ParseStyles(string stylesXml)
		{
			if (string.IsNullOrEmpty(stylesXml))
				return Array.Empty<Style>();

			List<Style> styles = null;
			var reader = XmlReader.Create(new StringReader(stylesXml));
			if (reader.ReadToFollowing("cellXfs"))
			{
				var subReader = reader.ReadSubtree();
				if (reader.MoveToAttribute("count"))
				{
					var count = int.Parse(reader.Value);
					styles = new List<Style>(count);
				}
				else
				{
					styles = new List<Style>(64);
				}
				while (subReader.ReadToFollowing("xf"))
				{
					if (subReader.MoveToAttribute("numFmtId"))
						styles.Add(new Style(int.Parse(subReader.Value)));
					else
						styles.Add(new Style(0));
				}
			}
			return styles?.ToArray() ?? Array.Empty<Style>();
		}
	}
}
