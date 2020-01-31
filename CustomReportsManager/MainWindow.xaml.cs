using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CustomReportsManager {
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window {
		public MainWindow() {
			InitializeComponent();

			DataContext = CustomReports.Configuration.Instance;
			DataGridComboBoxColumnFormat.ItemsSource = CustomReports.Configuration.SavingFormats;
			Closing += MainWindow_Closing;
			Loaded += (s, e) => {
				if (!CustomReports.Configuration.Instance.IsConfigReadedSuccessfull)
					MessageBox.Show(this, "Не удалось считать файл конфигурации: " +
						CustomReports.Configuration.Instance.ConfigFilePath + Environment.NewLine +
						"Создана новая конфигурация, заполненная стандартными значениями", 
						"Ошибка конфигурации", MessageBoxButton.OK, MessageBoxImage.Information);

				CustomReports.Configuration.Instance.ReportItems.CollectionChanged +=
					ReportItems_CollectionChanged;
			};
		}

		private void ReportItems_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e) {
			if (DataGridReports is null)
				return;

			Decorator border = VisualTreeHelper.GetChild(DataGridReports, 0) as Decorator;
			if (border != null) {
				ScrollViewer scrollViewer = border.Child as ScrollViewer;
				scrollViewer.ScrollToTop();
			}
		}

		private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
			if (!CustomReports.Configuration.Instance.IsNeedToSave)
				return;

			MessageBoxResult result = MessageBox.Show(
				this, "Имеются несохраненные изменения, хотите сохранить их перед выходом?", "Сохранение изменений",
				MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

			if (result == MessageBoxResult.Yes)
				CustomReports.Configuration.SaveConfiguration();
			else if (result == MessageBoxResult.Cancel)
				e.Cancel = true;
		}

		private void ButtonEditRecipients(object sender, RoutedEventArgs e) {
			try {
				string buttonTag = (sender as Button).Tag as string;

				WindowRecipientsListView windowRecipientsListView = 
					new WindowRecipientsListView(null, true) {
					Owner = this,
					Title = buttonTag
				};
				windowRecipientsListView.ShowDialog();
			} catch (Exception exc) {
				MessageBox.Show(this, exc.Message, string.Empty, 
					MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void ButtonRemoveSelected_Click(object sender, RoutedEventArgs e) {
			if (DataGridReports.SelectedItem == null) {
				MessageBox.Show(this, "Для удаления необходимо выбрать нужный отчет", "", 
					MessageBoxButton.OK, MessageBoxImage.Information);
				return;
			}

			MessageBoxResult result = MessageBox.Show(
				this,
				"Вы действительно хотите удалить выбранный отчет?",
				"",
				MessageBoxButton.YesNo,
				MessageBoxImage.Question);

			if (result == MessageBoxResult.No)
				return;

			CustomReports.Configuration.Instance.ReportItems.
				Remove(DataGridReports.SelectedItem as CustomReports.ItemReport);
		}

		private void ButtonCreate_Click(object sender, RoutedEventArgs e) {
			string errorMessage = string.Empty;

			if (DataGridReports.SelectedItem == null)
				errorMessage = "Не выбран отчет для формирования" + Environment.NewLine;

			if (string.IsNullOrEmpty(((DataGridReports.SelectedItem as CustomReports.ItemReport).SaveFormat)))
				errorMessage += "Не задан формат выгрузки" + Environment.NewLine;

			if (CustomReports.Configuration.Instance.DateBegin > 
				CustomReports.Configuration.Instance.DateEnd)
				errorMessage += "Дата окончания не быть может быть меньше даты начала";

			if (!string.IsNullOrEmpty(errorMessage)) {
				MessageBox.Show(this, errorMessage, "", MessageBoxButton.OK, MessageBoxImage.Warning);
				return;
			}

			CustomReports.ItemReport itemReport = 
				DataGridReports.SelectedItem as CustomReports.ItemReport;
			itemReport.SetPeriod(
				CustomReports.Configuration.Instance.DateBegin, 
				CustomReports.Configuration.Instance.DateEnd);
			string title = itemReport.Name + ", формирование за период с " +
				itemReport.DateBegin.ToShortDateString() + " по " + itemReport.DateEnd.ToShortDateString();

			WindowDetails windowDetails = new WindowDetails(title, string.Empty, this, itemReport);
			windowDetails.ShowDialog();
		}

		private void ButtonItemReport_Click(object sender, RoutedEventArgs e) {
			Button button = sender as Button;
			CustomReports.ItemReport itemReport = button.DataContext as CustomReports.ItemReport;

			if (itemReport == null) {
				MessageBox.Show(this, "ItemReport is null", "", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}

			string tag = button.Tag as string;
			if (string.IsNullOrEmpty(tag)) {
				MessageBox.Show(this, "Button.Tag is null", "", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}

			Window windowToShow = null;

			if (tag.Equals("CleanFolderToSave")) {
				itemReport.FolderToSave = string.Empty;
			} else if (tag.Equals("EditFolderToSave")) {
				CommonOpenFileDialog dlg = new CommonOpenFileDialog {
					Title = "Выбор пути сохранения",
					IsFolderPicker = true,
					AddToMostRecentlyUsedList = false,
					AllowNonFileSystemItems = false,
					EnsureFileExists = true,
					EnsurePathExists = true,
					EnsureReadOnly = false,
					EnsureValidNames = true,
					Multiselect = false,
					ShowPlacesList = true
				};

				if (dlg.ShowDialog() == CommonFileDialogResult.Ok) 
					itemReport.FolderToSave = dlg.FileName;
			} else if (tag.Equals("EditQuery")) {
				windowToShow = new WindowSqlQueryView(itemReport);
			} else if (tag.Equals("EditRecipients")) {
				windowToShow = new WindowRecipientsListView(itemReport);
			} else {
				MessageBox.Show(this, "Button.Tag unknown: " + tag, "", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}

			if (windowToShow != null) {
				windowToShow.Owner = this;
				windowToShow.ShowDialog();
			}
		}

		private void ButtonAddNewReport_Click(object sender, RoutedEventArgs e) {
			WindowAddNewReport windowAddNewReport = new WindowAddNewReport();
			windowAddNewReport.Owner = this;
			windowAddNewReport.ShowDialog();
		}
	}
}
