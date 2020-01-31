using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace CustomReports {
	[Serializable]
	public class Configuration : INotifyPropertyChanged {
		[field: NonSerialized]
		public event PropertyChangedEventHandler PropertyChanged;
		private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") {
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
			IsNeedToSave = true;
		}
		public bool IsConfigReadedSuccessfull { get; set; } = false;
		public bool IsNeedToSave { get; set; } = false;
		
		private string configFilePath;
		public string ConfigFilePath {
			get { return configFilePath; }
			set {
				if (value != configFilePath) {
					configFilePath = value;
					NotifyPropertyChanged();
				}
			}
		}

		#region DB
		private string misDbAddress;
		public string MisDbAddress {
			get { return misDbAddress; }
			set {
				if (value != misDbAddress) {
					misDbAddress = value;
					NotifyPropertyChanged();
				}
			}
		}
		private string misDbName;
		public string MisDbName {
			get { return misDbName; }
			set {
				if (value != misDbName) {
					misDbName = value;
					NotifyPropertyChanged();
				}
			}
		}
		private string misDbUserName;
		public string MisDbUserName {
			get { return misDbUserName; }
			set {
				if (value != misDbUserName) {
					misDbUserName = value;
					NotifyPropertyChanged();
				}
			}
		}
		private string misDbUserPassword;
		public string MisDbUserPassword {
			get { return misDbUserPassword; }
			set {
				if (value != misDbUserPassword) {
					misDbUserPassword = value;
					NotifyPropertyChanged();
				}
			}
		}
		#endregion


		#region Mail
		private string mailSmtpServer;
		public string MailSmtpServer {
			get { return mailSmtpServer; }
			set {
				if (value != mailSmtpServer) {
					mailSmtpServer = value;
					NotifyPropertyChanged();
				}
			}
		}
		private uint mailSmtpPort;
		public uint MailSmtpPort {
			get { return mailSmtpPort; }
			set {
				if (value != mailSmtpPort) {
					mailSmtpPort = value;
					NotifyPropertyChanged();
				}
			}
		}
		private string mailUser;
		public string MailUser {
			get { return mailUser; }
			set {
				if (value != mailUser) {
					mailUser = value;
					NotifyPropertyChanged();
				}
			}
		}
		private bool mailEnableSSL;
		public bool MailEnableSSL {
			get { return mailEnableSSL; }
			set {
				if (value != mailEnableSSL) {
					mailEnableSSL = value;
					NotifyPropertyChanged();
				}
			}
		}
		private string mailPassword;
		public string MailPassword {
			get { return mailPassword; }
			set {
				if (value != mailPassword) {
					mailPassword = value;
					NotifyPropertyChanged();
				}
			}
		}
		private string mailUserDomain;
		public string MailUserDomain {
			get { return mailUserDomain; }
			set {
				if (value != mailUserDomain) {
					mailUserDomain = value;
					NotifyPropertyChanged();
				}
			}
		}
		private string mailAdminAddress;
		public string MailAdminAddress {
			get { return mailAdminAddress; }
			set {
				if (value != mailAdminAddress) {
					mailAdminAddress = value;
					NotifyPropertyChanged();
				}
			}
		}
		private bool shouldAddAdminToCopy;
		public bool ShouldAddAdminToCopy {
			get { return shouldAddAdminToCopy; }
			set {
				if (value != shouldAddAdminToCopy) {
					shouldAddAdminToCopy = value;
					NotifyPropertyChanged();
				}
			}
		}

		private string mailSenderName;
		public string MailSenderName {
			get { return mailSenderName; }
			set {
				if (value != mailSenderName) {
					mailSenderName = value;
					NotifyPropertyChanged();
				}
			}
		}

		private string mailSign;
		public string MailSign {
			get { return mailSign; }
			set {
				if (value != mailSign) {
					mailSign = value;
					NotifyPropertyChanged();
				}
			}
		}
		#endregion


		#region General
		private uint maxLogfilesQuantity;
		public uint MaxLogfilesQuantity {
			get { return maxLogfilesQuantity; }
			set {
				if (value != maxLogfilesQuantity) {
					maxLogfilesQuantity = value;
					NotifyPropertyChanged();
				}
			}
		}
		#endregion



		private static Configuration instance = null;
		private static readonly object padlock = new object();

		public static Configuration Instance {
			get {
				lock (padlock) {
					if (instance == null)
						instance = LoadConfiguration();

					return instance;
				}
			}
		}

		[NonSerialized()] private ICommand buttonClick;

		[IgnoreDataMember]
		public ICommand ButtonClick {
			get {
				return buttonClick ??
					(buttonClick = new CommandHandler((object parameter) =>
					Action(parameter)));
			}
		}

		public ObservableCollection<ItemReport> ReportItems { get; private set; }

		public static ObservableCollection<string> SavingFormats { get; set; } = new ObservableCollection<string> { "Excel", "XML", "CSV" };

		private DateTime dateBegin;
		public DateTime DateBegin {
			get { return dateBegin; }
			set {
				if (value != dateBegin) {
					bool isNeedToSaveCurrentConfig = IsNeedToSave;
					dateBegin = value;
					NotifyPropertyChanged();
					IsNeedToSave = isNeedToSaveCurrentConfig;
				}
			}
		}
		private DateTime dateEnd;
		public DateTime DateEnd {
			get { return dateEnd; }
			set {
				if (value != dateEnd) {
					bool isNeedToSaveCurrentConfig = IsNeedToSave;
					dateEnd = value;
					NotifyPropertyChanged();
					IsNeedToSave = isNeedToSaveCurrentConfig;
				}
			}
		}


		public void Action(object parameter) {
			string param = parameter.ToString();

			if (param.Equals("CheckDbConnection")) {
				CheckDbConnection();
			} else if (param.Equals("SaveConfig")) {
				SaveConfiguration();
			} else if (param.Equals("CheckMailSettings")) {
				CheckMailServer();
			} else if (param.Equals("EquateEndDateToBeginDate")) {
				DateEnd = DateBegin;
			} else if (param.Equals("SetDatesToCurrentDay")) {
				DateBegin = DateTime.Now;
				DateEnd = DateBegin;
			} else if (param.Equals("SetDatesToCurrentWeek")) {
				DateEnd = DateTime.Now;
				int dayOfWeek = (int)DateEnd.DayOfWeek;
				if (dayOfWeek == 0)
					dayOfWeek = 7;
				DateBegin = DateEnd.AddDays(-1 * (dayOfWeek - 1));
			} else if (param.Equals("SetDatesToCurrentMonth")) {
				DateBegin = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
				DateEnd = DateBegin.AddDays(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month) - 1);
			} else if (param.Equals("SetDatesToCurrentYear")) {
				DateBegin = new DateTime(DateTime.Now.Year, 1, 1);
				DateEnd = new DateTime(DateTime.Now.Year, 12, DateTime.DaysInMonth(DateTime.Now.Year, 12));
			} else if (param.Equals("GoToPreviousMonth")) {
				DateEnd = new DateTime(DateBegin.Year, DateBegin.Month, 1).AddDays(-1);
				DateBegin = DateEnd.AddDays(-1 * (DateTime.DaysInMonth(DateEnd.Year, DateEnd.Month) - 1));
			} else if (param.Equals("GoToPreviousDay")) {
				DateBegin = DateBegin.AddDays(-1);
				DateEnd = DateBegin;
			} else if (param.Equals("GoToNextDay")) {
				DateBegin = DateBegin.AddDays(1);
				DateEnd = DateBegin;
			} else if (param.Equals("GoToNextMonth")) {
				DateBegin = new DateTime(DateBegin.Year, DateBegin.Month, DateTime.DaysInMonth(DateBegin.Year, DateBegin.Month)).AddDays(1);
				DateEnd = DateBegin.AddDays((DateTime.DaysInMonth(DateBegin.Year, DateBegin.Month) - 1));
			}
		}

		private async void CheckMailServer() {
			List<string> emptyStrings = new List<string>();

			if (string.IsNullOrEmpty(MailSmtpServer))
				emptyStrings.Add("Адрес SMTP-сервера");
			if (MailSmtpPort == 0)
				emptyStrings.Add("Порт");
			if (string.IsNullOrEmpty(MailUser))
				emptyStrings.Add("Имя пользователя");
			if (string.IsNullOrEmpty(MailPassword))
				emptyStrings.Add("Пароль");

			if (emptyStrings.Count > 0) {
				string msg = "Для проверки подключения к почтовому серверу необходимо заполнить:" + Environment.NewLine +
					string.Join(Environment.NewLine, emptyStrings);
				MessageBox.Show(msg, string.Empty, MessageBoxButton.OK, MessageBoxImage.Information);
				return;
			}

			if (string.IsNullOrEmpty(MailAdminAddress)) {
				MessageBox.Show("Необходимо задать как минимум один адрес в списке 'Получатели системных уведомлений'",
					string.Empty, MessageBoxButton.OK, MessageBoxImage.Information);
				return;
			}

			try {
				await Task.Run(() => {
					Mail.SendTestMail("Проверка подключения", "Проверка подключения", MailAdminAddress);
					MessageBox.Show("Проверка подключения: Успешно", string.Empty, MessageBoxButton.OK, MessageBoxImage.Information);
				}).ConfigureAwait(false);
			} catch (Exception e) {
				MessageBox.Show("Проверка подключения: Ошибка" +
					Environment.NewLine + e.Message, string.Empty, MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private async void CheckDbConnection() {
			List<string> emptyStrings = new List<string>();

			if (string.IsNullOrEmpty(MisDbAddress))
				emptyStrings.Add("Адрес БД МИС Инфоклиника");
			if (string.IsNullOrEmpty(MisDbName))
				emptyStrings.Add("Имя базы");
			if (string.IsNullOrEmpty(MisDbUserName))
				emptyStrings.Add("Имя пользователя");
			if (string.IsNullOrEmpty(MisDbUserPassword))
				emptyStrings.Add("Пароль");

			if (emptyStrings.Count > 0) {
				string msg = "Для проверки подключения к БД необходимо заполнить:" + Environment.NewLine +
					string.Join(Environment.NewLine, emptyStrings);
				MessageBox.Show(msg, string.Empty, MessageBoxButton.OK, MessageBoxImage.Information);
				return;
			}

			await Task.Run(() => {
				using (FirebirdClient firebirdClient = new FirebirdClient(
					MisDbAddress,
					MisDbName,
					MisDbUserName,
					MisDbUserPassword)) {
					DataTable dt = firebirdClient.GetDataTable("select date 'Now' from rdb$database");

					string msg = "Подключение к БД: ";
					MessageBoxImage image;
					if (dt.Rows.Count == 1) {
						msg += "Успешно";
						image = MessageBoxImage.Information;
					} else {
						msg += "Ошибка";
						image = MessageBoxImage.Error;
					}

					MessageBox.Show(msg, string.Empty, MessageBoxButton.OK, image);
				};
			}).ConfigureAwait(false);
		}


		private static Configuration LoadConfiguration() {
			Configuration configuration = new Configuration();
			Logging.ToLog("Configuration - Считывание файла настроек: " + configuration.configFilePath);

			if (!File.Exists(configuration.configFilePath)) {
				Logging.ToLog("Configuration - !!! Не удается найти файл");
				return configuration;
			}

			try {

				byte[] key = { 1, 2, 3, 4, 5, 6, 7, 8 };
				byte[] iv = { 1, 2, 3, 4, 5, 6, 7, 8 };

				using (DESCryptoServiceProvider des = new DESCryptoServiceProvider())
				using (var fs = new FileStream(configuration.ConfigFilePath, FileMode.Open, FileAccess.Read))
				using (var cryptoStream = new CryptoStream(fs, des.CreateDecryptor(key, iv), CryptoStreamMode.Read)) {
					BinaryFormatter formatter = new BinaryFormatter();
					configuration = (Configuration)formatter.Deserialize(cryptoStream);
					configuration.IsConfigReadedSuccessfull = true;

					foreach (ItemReport item in configuration.ReportItems) {
						if (string.IsNullOrEmpty(item.SaveFormat))
							item.SaveFormat = SavingFormats[0];
					}
				};
			} catch (Exception e) {
				Logging.ToLog("Configuration - !!! " + e.Message + Environment.NewLine + e.StackTrace);
			}

			configuration.DateBegin = DateTime.Now;
			configuration.DateEnd = DateTime.Now;
			configuration.IsNeedToSave = false;

			return configuration;
		}

		public static bool SaveConfiguration() {
			byte[] key = { 1, 2, 3, 4, 5, 6, 7, 8 };
			byte[] iv = { 1, 2, 3, 4, 5, 6, 7, 8 };
			try {
				using (DESCryptoServiceProvider des = new DESCryptoServiceProvider())
				using (var fs = new FileStream(Instance.ConfigFilePath, FileMode.Create, FileAccess.Write))
				using (var cryptoStream = new CryptoStream(fs, des.CreateEncryptor(key, iv), CryptoStreamMode.Write)) {
					BinaryFormatter formatter = new BinaryFormatter();
					formatter.Serialize(cryptoStream, Instance);
				};

				Instance.IsNeedToSave = false;
				MessageBox.Show("Изменения сохранены", string.Empty,
					MessageBoxButton.OK, MessageBoxImage.Information);

				return true;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				MessageBox.Show("Не удалось сохранить конфигурацию: " + e.Message,
					"Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

				return false;
			}
		}

		private Configuration() {
			ConfigFilePath = Logging.AssemblyDirectory + "CustomReports.cfg";
			ReportItems = new ObservableCollection<ItemReport>();
			ReportItems.CollectionChanged += (s, e) => {
				IsNeedToSave = true;
			};

			LoadDefaults();
		}

		public static string[] GetSplittedAddresses(string addresses) {
			return addresses.Split(new string[] { " | " }, StringSplitOptions.RemoveEmptyEntries);
		}

		private void LoadDefaults() {
			#region DB
			MisDbAddress = "127.0.0.1";
			MisDbName = "db_name";
			MisDbUserName = "sysdba";
			MisDbUserPassword = "masterkey";
			#endregion

			#region Mail
			MailSmtpServer = "smtp.server.ru";
			MailSmtpPort = 587;
			MailUser = "donotreply@server.ru";
			MailEnableSSL = false;
			MailPassword = string.Empty;
			MailUserDomain = string.Empty;
			MailAdminAddress = string.Empty;
			ShouldAddAdminToCopy = true;
			MailSenderName = "CustomReportsManager";
			MailSign = "___________________________________________" + Environment.NewLine +
				"Это автоматически сгенерированное сообщение" + Environment.NewLine +
				"Просьба не отвечать на него" + Environment.NewLine +
				 "Имя системы: @machineName";
			#endregion

			#region General
			MaxLogfilesQuantity = 14;
			#endregion
		}

		public static ItemReport GetReportByID(string id) {
			ItemReport item = null;

			foreach (ItemReport report in Instance.ReportItems)
				if (report.ID.ToUpper().Equals(id.ToUpper()))
					item = report;

			return item;
		}
	}
}
