using System.Collections.ObjectModel;

namespace XLSXInfo.Model
{
	public class WorkBook
	{
		public ObservableCollection<Worksheet> Worksheets { get; set; }
	}
}
