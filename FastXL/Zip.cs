using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;

namespace FastXL
{
	internal static class ZipArchiveEntryExtension
	{
		public static string ReadAllText(this ZipArchiveEntry entry)
		{
			using (var stream = entry.Open())
			using (var reader = new StreamReader(stream, Encoding.UTF8))
				return reader.ReadToEnd();
		}

		public static async Task<string> ReadAllTextAsync(this ZipArchiveEntry entry)
		{
			using (var stream = entry.Open())
			using (var reader = new StreamReader(stream, Encoding.UTF8))
				return await reader.ReadToEndAsync();
		}
	}

	internal class Zip : IDisposable
	{
		Stream stream;
		ZipArchive zipArchive;

		public IReadOnlyList<ZipArchiveEntry> Entries => zipArchive.Entries;

		public static async Task<Zip> LoadAsync(Stream stream)
		{
			byte[] buffer;
			if (stream is MemoryStream memoryStream)
			{
				buffer = memoryStream.ToArray();
				memoryStream.Dispose();
			}
			else
			{
				buffer = new byte[(int)stream.Length];
				await stream.ReadAsync(buffer, 0, buffer.Length);
				stream.Close();
				stream.Dispose();
			}

			return new Zip(new MemoryStream(buffer, false));
		}

		Zip(Stream stream)
		{
			this.stream = stream ?? throw new ArgumentNullException(nameof(stream));
			zipArchive = new ZipArchive(this.stream, ZipArchiveMode.Read, false);
		}

		~Zip() => Dispose();

		public void Dispose()
		{
			zipArchive?.Dispose();
			zipArchive = null;
			stream?.Dispose();
			stream = null;
		}
	}
}
