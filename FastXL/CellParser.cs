using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace FastXL
{
	static class CellParser
	{
		static readonly object trueValue = true;
		static readonly object falseValue = false;
		static readonly object intZero = 0;
		static readonly object longZero = 0L;
		static readonly object ulongZero = 0UL;
		static readonly object doubleZero = 0d;
		static readonly object decimalZero = 0m;

		static readonly Regex floatRE = new Regex(@"^-?\d+(?:\.?\d*[eE][+-]?|\.)\d+$", RegexOptions.Compiled | RegexOptions.Multiline);
		static readonly Regex intRE = new Regex(@"^-?\d+(?:[eE]\+?\d+)?$", RegexOptions.Compiled | RegexOptions.Multiline);

		public static object Parse(string type, string style, string value, WorkbookContext context)
		{
			if (value == null)
				return null;

			if (TryParseByAttribute(type, style, value, context, out var parsedValue))
				return parsedValue;

			if (floatRE.IsMatch(value))
			{
				if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var doubleValue)) return doubleValue == 0.0d ? doubleZero : doubleValue;
				if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var decimalValue)) return decimalValue == 0.0m ? decimalZero : decimalValue;
				return value;
			}

			if (intRE.IsMatch(value))
			{
				if (int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var intValue)) return intValue == 0 ? intZero : intValue;
				if (long.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var longValue)) return longValue == 0L ? longZero : longValue;
				if (ulong.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var ulongValue)) return ulongValue == 0UL ? ulongZero : ulongValue;
			}

			return value;
		}

		static bool TryParseByAttribute(string type, string style, string value, WorkbookContext context, out object parsedValue)
		{
			parsedValue = null;
			if (!string.IsNullOrEmpty(type))
			{
				switch (type)
				{
					// t="s" 는 값 문자열이 공유 문자열 인덱스이다.
					case "s":
						parsedValue = context.SharedStrings[int.Parse(value)];
						return true;

					// t="b" 는 값 문자열을 논리값으로 해석한다. (0: false, 1:true)
					case "b":
						parsedValue = value != "0" ? trueValue : falseValue;
						return true;
				}
			}

			// s="1" 에서 숫자는 공유 스타일 인덱스이다.
			if (!string.IsNullOrEmpty(style))
			{
				var styleIndex = int.Parse(style);
				var styleObj = context.Styles[styleIndex];
				switch (styleObj.NumberFormat)
				{
					// 스타일의 숫자형식(numFmtId)이 DateTime(14)인 것은 날짜로 해석한다.
					case NumberFormat.DateTime:
						{
							var dateTimeValue = double.Parse(value);
							parsedValue = DateTime.FromOADate(dateTimeValue);
							return true;
						}
				}
			}

			return false;
		}
	}
}
