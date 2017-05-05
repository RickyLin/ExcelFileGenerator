namespace RL.FileGenerators.Excel
{
    public class DefinedNameData
	{
		public string Name { get; set; }
		public string SheetName { get; set; }
		public string Reference { get; set; }
		public string ColumnName { get; set; }
		public uint RowIndex { get; set; }

		public DefinedNameData() { }

		public DefinedNameData(string name, string sheetName, string columnName, uint rowIndex)
		{
			Name = name;
			SheetName = sheetName;
			ColumnName = columnName;
			RowIndex = rowIndex;
			Reference = columnName + rowIndex.ToString();
		}
	}
}
