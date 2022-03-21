namespace ExcelLightMapper
{
	using System;

	public class ExCell
	{
		public string Sheet { get; set; }
		public string SheetCodeName { get; set; }

		public int RowIndex { get; set; }
		public int Row => RowIndex + 1;

		public int? ColumnIndex { get; set; }
		public int? Column => ColumnIndex.HasValue ? ColumnIndex.Value + 1 : ColumnIndex;

		public ExCell()
		{
		}

		public ExCell(int rowIndex, int? columnIndex)
		{
			RowIndex = rowIndex;
			ColumnIndex = columnIndex;
		}

		public ExCell(ExCell cell)
		{
			Sheet = cell.Sheet;
			SheetCodeName = cell.SheetCodeName;
			RowIndex = cell.RowIndex;
			ColumnIndex = cell.ColumnIndex;
		}

		public override string ToString()
		{
			return $"[{Sheet}]:({Column},{Row})";
		}

		public override bool Equals(object obj)
		{
			if (obj is ExCell other)
			{
				return Sheet.Equals(other.Sheet, StringComparison.InvariantCultureIgnoreCase)
					&& RowIndex == other.RowIndex
					&& ColumnIndex == other.ColumnIndex;
			}
			return false;
		}
	}
}