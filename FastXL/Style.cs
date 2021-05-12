namespace FastXL
{
	enum NumberFormat
	{
		Normal = 0,
		Percent = 9,
		DateTime = 14
	}

	readonly struct Style
	{
		public readonly NumberFormat NumberFormat;

		public Style(int numberFormat)
		{
			NumberFormat = (NumberFormat)numberFormat;
		}
	}
}
