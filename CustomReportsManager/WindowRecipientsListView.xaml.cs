using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
	/// Interaction logic for WindowRecipientsListView.xaml
	/// </summary>
	public partial class WindowRecipientsListView : Window {
		public ObservableCollection<MailAddress> Addresses { get; set; } = new ObservableCollection<MailAddress>();

		public WindowRecipientsListView(CustomReports.ItemReport itemReport, bool isAdminAddress = false) {
			InitializeComponent();

			string addresses = string.Empty;
			if (isAdminAddress) {
				addresses = CustomReports.Configuration.Instance.MailAdminAddress;
				Title = "Редактирования списка получателей системных уведомлений";
			} else if (itemReport == null) {
				MessageBox.Show(this, "ItemReport is null", "", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			} else {
				addresses = itemReport.Recipients;
				Title = "Редактирование списка получателей для: '" + itemReport.Name + "'";
			}

			string[] splitted = CustomReports.Configuration.GetSplittedAddresses(addresses);
			foreach (string address in splitted)
				Addresses.Add(new MailAddress(address));

			DataGridAddresses.DataContext = this;

			Closed += (s, e) => {
				List<MailAddress> emptyOrWrong = new List<MailAddress>();
				foreach (MailAddress item in Addresses) {
					if (string.IsNullOrEmpty(item.Address)) {
						emptyOrWrong.Add(item);
						continue;
					}

					try {
						System.Net.Mail.MailAddress mailAddress = new System.Net.Mail.MailAddress(item.Address);
					} catch (Exception) {
						emptyOrWrong.Add(item);
					}
				}

				foreach (MailAddress item in emptyOrWrong)
					Addresses.Remove(item);

				string addressesEdited = string.Join(" | ", Addresses);

				if (isAdminAddress)
					CustomReports.Configuration.Instance.MailAdminAddress = addressesEdited;
				else
					itemReport.Recipients = addressesEdited;
			};
		}

		public class MailAddress {
			public string Address { get; set; }
			public MailAddress(string address) {
				Address = address;
			}

			override
			public string ToString() {
				return Address;
			}
		}

		private void ButtonAdd_Click(object sender, RoutedEventArgs e) {
			Addresses.Add(new MailAddress(string.Empty));

			DataGridAddresses.SelectedIndex = Addresses.Count - 1;
			DataGridAddresses.Focus();
		}

		private void DataGridAddresses_SelectionChanged(object sender, SelectionChangedEventArgs e) {
			ButtonRemove.IsEnabled = DataGridAddresses.SelectedItems.Count > 0;
		}

		private void ButtonRemove_Click(object sender, RoutedEventArgs e) {
			List<MailAddress> mailAddresses = new List<MailAddress>();
			foreach (MailAddress mailAddress in DataGridAddresses.SelectedItems)
				mailAddresses.Add(mailAddress);

			foreach (MailAddress mailAddressToRemove in mailAddresses)
				Addresses.Remove(mailAddressToRemove);
		}

		private void ButtonClose_Click(object sender, RoutedEventArgs e) {
			Close();
		}
	}
}
