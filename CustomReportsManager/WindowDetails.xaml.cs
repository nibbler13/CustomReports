using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
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
	/// Interaction logic for WindowDetails.xaml
	/// </summary>
	public partial class WindowDetails : Window {
		public WindowDetails(string title, string text, Window owner, CustomReports.ItemReport itemReport = null) {
			InitializeComponent();
			Title = title;
			TextBoxMain.Text = text;
			Owner = owner;

			if (itemReport != null) {
				CreateReport(itemReport);
				ButtonClose.IsEnabled = false;
				Cursor = Cursors.Wait;
			}
		}

		private void ButtonClose_Click(object sender, RoutedEventArgs e) {
			Close();
		}

		private void CreateReport(CustomReports.ItemReport itemReport) {
			BackgroundWorker bw = new BackgroundWorker();
			bw.WorkerReportsProgress = true;

			bw.ProgressChanged += (s, e) => {
				if (e.UserState != null) {
					TextBoxMain.Text += e.UserState.ToString() + Environment.NewLine;
					TextBoxMain.ScrollToEnd();
				}
			};

			bw.DoWork += (s, e) => {
				CustomReports.Program.CreateReport(itemReport);
			};

			bw.RunWorkerCompleted += (s, e) => {
				if (e.Error != null) {
					MessageBox.Show(this, e.Error.Message + Environment.NewLine + e.Error.StackTrace, "",
						MessageBoxButton.OK, MessageBoxImage.Error);
				} else {
					MessageBox.Show(this, "Завершено", "",
						MessageBoxButton.OK, MessageBoxImage.Information);
				}

				ButtonClose.IsEnabled = true;
				Cursor = Cursors.Arrow;

				if (!string.IsNullOrEmpty(itemReport.FileResult)) {
					try {
						string argument = "/select, \"" + itemReport.FileResult + "\"";
						Process.Start("explorer.exe", argument);
					} catch (Exception exc) {
						MessageBox.Show(this, exc.Message, string.Empty, MessageBoxButton.OK, MessageBoxImage.Error);
					}
				}
			};

			CustomReports.Logging.bw = bw;
			bw.RunWorkerAsync();
		}
	}
}
