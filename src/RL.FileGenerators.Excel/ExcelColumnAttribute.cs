using System;

namespace RL.FileGenerators.Excel
{
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
	public sealed class ExcelColumnAttribute : Attribute
	{
		/// <summary>
		///  the column title
		/// </summary>
		public string Title { get; private set; }

		/// <summary>
		/// the column index in the excel, started from 1
		/// </summary>
		public int Index { get; private set; }

		/// <summary>
		/// the Excel embedded number format index used to format the value
		/// </summary>
		public uint NumberFormatId { get; set; }

		/// <summary>
		/// customized number format string to format the value
		/// </summary>
		public string NumberingFormatString { get; set; }

		public ExcelColumnAttribute(string title, int index)
		{
			Title = title;
			Index = index;
			NumberFormatId = uint.MaxValue;
		}
	}
}
