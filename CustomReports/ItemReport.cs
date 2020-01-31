using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace CustomReports {
	[Serializable]
	public class ItemReport : INotifyPropertyChanged {
		[field: NonSerialized]
		public event PropertyChangedEventHandler PropertyChanged;
		private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") {
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
			Configuration.Instance.IsNeedToSave = true;
		}


		public string SaveFormat { get; set; } = "Excel";


		private string id;
		public string ID {
			get { return id; }
			private set {
				if (value != id) {
					id = value;
					NotifyPropertyChanged();
				}
			}
		}

		private string name;
		public string Name {
			get { return name; }
			set {
				if (value != name) {
					name = value;
					NotifyPropertyChanged();
				}
			}
		}

		private string recipients;
		public string Recipients {
			get { return recipients; }
			set {
				if (value != recipients) {
					recipients = value;
					NotifyPropertyChanged();
					NotifyPropertyChanged(nameof(RecipientsCount));
				}
			}
		}

		public string RecipientsCount {
			get {
				if (string.IsNullOrEmpty(Recipients))
					return "Не заданы";
				else
					return "Кол-во: " + Configuration.GetSplittedAddresses(Recipients).Length; 
			}
		}

		private string query;
		public string Query {
			get { return query; }
			set {
				if (value != query) {
					query = value;
					NotifyPropertyChanged();
					NotifyPropertyChanged("QueryCount");
				}
			}
		}

		public string QueryCount {
			get {
				if (Query.Length == 0)
					return "Не задан";
				else
					return "Кол-во строк: " + Query.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Length;
			}
		}

		private string folderToSave;
		public string FolderToSave {
			get { return folderToSave; }
			set {
				if (value != folderToSave) {
					folderToSave = value;
					NotifyPropertyChanged();
				}
			}
		}

		private bool shouldBeSavedToFolder;
		public bool ShouldBeSavedToFolder {
			get { return shouldBeSavedToFolder; }
			set {
				if (value != shouldBeSavedToFolder) {
					shouldBeSavedToFolder = value;
					NotifyPropertyChanged();
				}
			}
		}

		public DateTime DateBegin { get; set; }
		public DateTime DateEnd { get; set; }
		public void SetPeriod(DateTime dateBegin, DateTime dateEnd) {
			DateBegin = dateBegin;
			DateEnd = dateEnd;
		}

		public string FileResult { get; set; }

		public ItemReport(string id) {
			ID = id;
			Name = "Введите название";
			Recipients = string.Empty;
			Query = string.Empty;
			FolderToSave = string.Empty;
			ShouldBeSavedToFolder = false;
		}
	}
}
