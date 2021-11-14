using System;
using System.Diagnostics;

namespace AppAsToy.FastXL
{
	[DebuggerDisplay("{Value}")]
	public readonly struct Cell
	{
		static readonly string defaultDateTimeFormat = @"yyyy\-MM\-dd";

		public readonly object Value;

		public Cell(object value)
		{
			Value = value;
		}

		public override string ToString()
		{
			if (Value is DateTime datetime)
				return datetime.ToString(defaultDateTimeFormat);
			return Value?.ToString() ?? string.Empty;
		}

		public string ToString(string dateTimeFormat)
		{
			if (Value is DateTime datetime)
				return datetime.ToString(dateTimeFormat);
			return Value?.ToString() ?? string.Empty;
		}
	}
}
