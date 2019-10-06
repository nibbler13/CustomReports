using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace CustomReportsManager {
	/// <summary>
	/// Interaction logic for WindowAddNewReport.xaml
	/// </summary>
	public partial class WindowAddNewReport : Window, INotifyPropertyChanged {
		[field: NonSerialized]
		public event PropertyChangedEventHandler PropertyChanged;
		private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") {
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}

		private Regex regex = new Regex("[^A-Za-z0-9]");

		private string enteredValue;
		public string EnteredValue {
			get { return enteredValue; }
			set {
				if (value != enteredValue) {
					enteredValue = value;

					bool isExistSomeReportWithSameID = false;
					
					if (!string.IsNullOrEmpty(enteredValue)) {
						foreach (CustomReports.ItemReport item in CustomReports.Configuration.Instance.ReportItems) {
							if (string.IsNullOrEmpty(item.ID))
								continue;

							if (item.ID.ToUpper().Equals(enteredValue.ToUpper())) {
								isExistSomeReportWithSameID = true;
								break;
							}
						}
					}

					if (string.IsNullOrEmpty(enteredValue)) {
						TextHint = "Индентификатор не должен быть пустым";
						ButtonOK.IsEnabled = false;
					} else if (regex.IsMatch(enteredValue)) {
						TextHint = "Идентификатор не должен содержать спецсимволы или кириллицу";
						ButtonOK.IsEnabled = false;
					} else if (isExistSomeReportWithSameID) {
						TextHint = "Отчет с таким ID уже сущесвует";
						ButtonOK.IsEnabled = false;
					} else {
						TextHint = string.Empty;
						ButtonOK.IsEnabled = true;
					}

					NotifyPropertyChanged();
				}
			}
		}

		private string textAbout;
		public string TextAbout {
			get { return textAbout; }
			set {
				if (value != textAbout) {
					textAbout = value;
					NotifyPropertyChanged();
				}
			}
		}
		private string textHint;
		public string TextHint {
			get { return textHint; }
			set {
				if (value != textHint) {
					textHint = value;
					NotifyPropertyChanged();
				}
			}
		}

		public WindowAddNewReport() {
			InitializeComponent();
			DataContext = this;

			EnteredValue = string.Empty;
			TextAbout = "Для каждого отчета должен быть указан уникальный идентификатор." + Environment.NewLine +
				"Он должен состоять из латинских букв (регистр не важен) и / или цифр, но не должен содержать пробелов." + Environment.NewLine +
				"Идентификатор используется для запуска запланированного задания в системном планировщике заданий.";

			Loaded += (s, e) => { TextBoxID.Focus(); };
		}

		private void ButtonOK_Click(object sender, RoutedEventArgs e) {
			CustomReports.Configuration.Instance.ReportItems.Insert(
				0, new CustomReports.ItemReport(EnteredValue));

			MessageBox.Show(
				this,
				"Новый отчет добавлен. Заполните информацию в соответствующей " +
				"строке, после чего сохраните изменения.", "",
				MessageBoxButton.OK, MessageBoxImage.Information);

			Close();
		}

		private void TextBox_PreviewKeyDown(object sender, KeyEventArgs e) {
			if (e.Key == Key.Space) {
				e.Handled = true;
			} else if (e.Key == Key.Enter && ButtonOK.IsEnabled) {
				ButtonOK_Click(null, null);
				e.Handled = true;
			}
		}
	}
}
