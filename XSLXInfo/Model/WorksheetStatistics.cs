namespace XLSXInfo.Model
{
	public class WorksheetStatistics
	{
		public int StartRowIndex { get; set; }
		public int EndRowIndex { get; set; }
		public int StartColumnIndex { get; set; }
		public int EndColumnIndex { get; set; }
		public int NumberOfColumns { get; set; }
		public int NumberOfRows { get; set; }
		public int NumberOfCells { get; set; }
		public int NumberOfEmptyCells { get; set; }
	}
}