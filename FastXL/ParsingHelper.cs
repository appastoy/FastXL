using System.Text.RegularExpressions;

namespace AppAsToy.FastXL
{
	static class ParsingHelper
	{
		const int alphabetRange = 'Z' - 'A' + 1;
		static readonly Regex addressRE = new Regex(@"([a-zA-Z]+)(\d*)", RegexOptions.Compiled);

		public static (int Row, int Column) ParseAddress(string address)
		{
			var match = addressRE.Match(address);
			var column = ConvertToNumber(match.Groups[1].Value);
			if (!int.TryParse(match.Groups[2].Value, out var row))
				row = 0;
			return (row, column);
		}

		public static int ConvertToNumber(string address)
		{
			var sum = 0;
			var unit = 1;
			for (int i = address.Length - 1; i >= 0; i--)
			{
				var ch = char.ToUpper(address[i]);
				sum += (ch - 'A' + 1) * unit;
				unit *= alphabetRange;
			}
			return sum;
		}
	}
}
