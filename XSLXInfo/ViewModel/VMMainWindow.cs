using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using SpreadsheetLight;
using XLSXInfo.Model;

namespace XLSXInfo.ViewModel
{
	public class VMMainWindow : INotifyPropertyChanged
	{
		private string filePath = String.Empty;
		private bool isShowProgressRing;
		
		public WorkBook WorkBook { get; set; }
		public ObservableCollection<Worksheet> SelectedSheetList { get; set; }
		public string FilePath
		{
			get => filePath;
		}
		public string Measurement { get; set; }
		public bool IsShowProgressRing
        {
			get => isShowProgressRing;
            set
            {
				isShowProgressRing = value;
				OnPropertyChanged(nameof(IsShowProgressRing));
			}
        }
		public bool IsEnabledContent => !isShowProgressRing;

		public string FileSize
		{
			get
			{
				if (string.IsNullOrEmpty(filePath))
				{
					return string.Empty;
				}

				var size = new FileInfo(filePath).Length;

				if (size < 1048576)
				{
					Measurement = "Kbite";
					return (size / 1024m).ToString(".##");
				}			
			
				Measurement = "Mbite";
				return (size / 1024m / 1024m).ToString(".##");
			}
		}

		public Worksheet SelectedSheet
		{
			set
			{
				SelectedSheetList[0] = value;
				OnPropertyChanged(nameof(SelectedSheet));
			}
		}
	
		public VMMainWindow()
		{
			WorkBook = new WorkBook();
			WorkBook.Worksheets = new ObservableCollection<Worksheet>();
			SelectedSheetList = new ObservableCollection<Worksheet>();
			SelectedSheetList.Add(new Worksheet());
		}

		public ICommand PressSelectFile
		{
			get => new ManagerCommand(SelectFile);
		}

		private void SelectFile()
		{
			var openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Open XML Spreadsheet (*.xlsx)|*.xlsx";
			if (openFileDialog.ShowDialog() == true)
			{
				filePath = openFileDialog.FileName;
				WorkBook.Worksheets.Clear();
				Task.Run(() =>
                {
                    try
                    {
						IsShowProgressRing = true;
						GetInfo(filePath);
					}
                    finally
                    {
						IsShowProgressRing = false;
                    }
                });
				if (WorkBook.Worksheets.Count > 0)
				{
					SelectedSheetList[0] = WorkBook.Worksheets[0];
				}
				OnPropertyChanged(nameof(FilePath));
				OnPropertyChanged(nameof(FileSize));
				OnPropertyChanged(nameof(Measurement));
			}
		}

		private void GetInfo(string file)
		{
			using (var document = new SLDocument(file))
			{
				foreach (var sheet in document.GetWorksheetNames())
				{
					document.SelectWorksheet(sheet);
					var statistic = document.GetWorksheetStatistics();

					Application.Current.Dispatcher.Invoke(() => WorkBook.Worksheets.Add
					(
						new Worksheet()
						{
							Index = WorkBook.Worksheets.Count + 1,
							Name = sheet,
							Statistics = new WorksheetStatistics()
							{
								StartRowIndex = statistic.StartRowIndex,
								EndRowIndex = statistic.EndRowIndex,
								StartColumnIndex = statistic.StartColumnIndex,
								EndColumnIndex = statistic.EndColumnIndex,
								NumberOfColumns = statistic.NumberOfColumns,
								NumberOfRows = statistic.NumberOfRows,
								NumberOfCells = statistic.NumberOfCells,
								NumberOfEmptyCells = statistic.NumberOfEmptyCells
							}
						}
					));
				}
			}
		}

		public event PropertyChangedEventHandler PropertyChanged;

		protected virtual void OnPropertyChanged(string propertyName)
		{
			if (propertyName == nameof(IsShowProgressRing))
            {
				PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsEnabledContent)));
			}

			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}
	}
}