using System.Reflection;

namespace RL.FileGenerators.Excel
{
    public class ExcelColumnProperty
	{
		public PropertyInfo PropertyInfo { get; set; }
		public ExcelColumnAttribute ExcelColumnAttr { get; set; }
		public int CellFormatIndex { get; set; }

		public ExcelColumnProperty()
		{
			CellFormatIndex = -1;
		}
	}
}
