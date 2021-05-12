using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace FastXL
{
	internal class Zip : IDisposable
	{
		readonly ZipArchive zipArchive;
		bool isDisposed;

		public IReadOnlyList<ZipArchiveEntry> Entries => zipArchive.Entries;

		public static Zip Open(Stream stream)
		{
			return new Zip(stream);
		}

		Zip(Stream stream)
		{
			zipArchive = new ZipArchive(stream, ZipArchiveMode.Read, false);
		}

		public string ReadString(string fileName)
		{
			var entry = zipArchive.Entries.FirstOrDefault(e => e.Name == fileName);
			if (entry == null)
				return string.Empty;

			return ReadAllText(entry);
		}

		static string ReadAllText(ZipArchiveEntry entry)
		{
			using (var stream = entry.Open())
			using (var reader = new StreamReader(stream))
				return reader.ReadToEnd();
		}

		public void Dispose()
		{
			if (!isDisposed)
			{
				zipArchive.Dispose();
				GC.SuppressFinalize(this);
				isDisposed = true;
			}
		}
	}
}
