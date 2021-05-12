using System;
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

		static readonly Regex valueRE = new Regex(@"<v>(.+)</v>", RegexOptions.Compiled);
		static readonly Regex floatRE = new Regex(@"^-?\d+(?:\.?[eE][+-]?|\.)\d+$", RegexOptions.Compiled | RegexOptions.Multiline);
		static readonly Regex intRE = new Regex(@"^-?\d+(?:[eE]\+?\d+)?$", RegexOptions.Compiled | RegexOptions.Multiline);
		static readonly Regex typeRE = new Regex(@"\bt=""(\w+)""");
		static readonly Regex styleRE = new Regex(@"\bs=""(\d+)""");

		public static object Parse(string attrXml, string valueXml, WorkbookContext context)
		{
			var match = valueRE.Match(valueXml);
			if (!match.Success)
				return valueXml;

			var valueString = match.Groups[1].Value;
			if (TryParseByAttribute(attrXml, valueString, context, out var parsedValue))
				return parsedValue;

			if (floatRE.IsMatch(valueString))
			{
				if (double.TryParse(valueString, out var doubleValue)) return doubleValue == 0.0 ? doubleZero : doubleValue;
				if (decimal.TryParse(valueString, out var decimalValue)) return decimalValue == 0m ? decimalZero : decimalValue;
				return valueString;
			}

			if (intRE.IsMatch(valueString))
			{
				if (int.TryParse(valueString, out var intValue)) return intValue == 0 ? intZero : intValue;
				if (long.TryParse(valueString, out var longValue)) return longValue == 0L ? longZero : longValue;
				if (ulong.TryParse(valueString, out var ulongValue)) return ulongValue == 0UL ? ulongZero : ulongValue;
			}

			return valueString;
		}

		/// <summary>
		/// 속성으로 파싱을 시도한다.
		/// <para>1. t="s" 는 값 문자열이 공유 문자열 인덱스이다.</para>
		/// <para>2. t="b" 는 값 문자열을 논리값으로 해석한다. (0: false, 1:true)</para>
		/// <para>3. s="1" 에서 숫자는 공유 스타일 인덱스인데, 스타일 숫자형식값이 DateTime(14)일 경우, 날짜로 해석한다.</para>
		/// </summary>
		/// <param name="attrXml">속성</param>
		/// <param name="valueString">값 문자열</param>
		/// <param name="context">엑셀 컨텍스트</param>
		/// <param name="parsedValue">파싱된 결과값</param>
		/// <returns>파싱 성공 여부</returns>
		static bool TryParseByAttribute(string attrXml, string valueString, WorkbookContext context, out object parsedValue)
		{
			parsedValue = null;
			var typeMatch = typeRE.Match(attrXml);
			if (typeMatch.Success)
			{
				var typeString = typeMatch.Groups[1].Value;
				switch (typeString)
				{
					case "s":
						parsedValue = context.SharedStrings[int.Parse(valueString)];
						return true;
					case "b":
						parsedValue = valueString != "0" ? trueValue : falseValue;
						return true;
				}
			}

			var styleMatch = styleRE.Match(attrXml);
			if (styleMatch.Success)
			{
				var styleIndex = int.Parse(styleMatch.Groups[1].Value);
				var style = context.Styles[styleIndex];
				switch (style.NumberFormat)
				{
					case NumberFormat.DateTime:
						{
							var dateTimeValue = double.Parse(valueString);
							parsedValue = DateTime.FromOADate(dateTimeValue);
							return true;
						}
				}
			}

			return false;
		}
	}
}
