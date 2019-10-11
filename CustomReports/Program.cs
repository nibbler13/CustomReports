using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CustomReports {
	public static class Program {
		public static string AssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";

		private static ItemReport itemReport;
		private static DataTable dataTableMainData = null;
		private static Dictionary<string, object> parameters;
		private static string dateBeginStr = string.Empty;
		private static string dateEndStr = string.Empty;
		private static string subject = string.Empty;
		private static string body = string.Empty;
		private static bool hasError = false;

		public static void Main(string[] args) {
			Logging.ToLog("Старт");

			if (!Configuration.Instance.IsConfigReadedSuccessfull) {
				Logging.ToLog("Отсутсвует файл конфигурации. " +
					"Воспользуйтесь утилитой CustomReportsManager.exe " +
					"для создания файла конфигурации.");
				return;
			}

			if (args.Length < 2 || args.Length > 3) {
				Logging.ToLog("Неверное количество параметров");
				WriteOutAcceptedParameters();
				return;
			}

			string reportID = args[0];
			itemReport = Configuration.GetReportByID(reportID);
			if (itemReport == null) {
				Logging.ToLog("Неизвестный ID отчета: " + reportID);
				WriteOutAcceptedParameters();
				return;
			}

			ParseDateInterval(args);

			if (itemReport.DateBegin == null || itemReport.DateEnd == null) {
				Logging.ToLog("Не удалось распознать временные интервалы формирования отчета");
				WriteOutAcceptedParameters();
				return;
			}

			CreateReport(itemReport);
		}

		public static void CreateReport(ItemReport itemReportToCreate) {
			itemReport = itemReportToCreate;
			dateBeginStr = itemReport.DateBegin.ToShortDateString();
			dateEndStr = itemReport.DateEnd.ToShortDateString();
			subject = itemReport.Name + " с " + dateBeginStr + " по " + dateEndStr;
			Logging.ToLog("Формирование: " + subject);

			using (FirebirdClient firebirdClient = new FirebirdClient(
				Configuration.Instance.MisDbAddress,
				Configuration.Instance.MisDbName,
				Configuration.Instance.MisDbUserName,
				Configuration.Instance.MisDbUserPassword)) {
				LoadData(firebirdClient);
			}

			WriteDataToFile();

			if (hasError) {
				Logging.ToLog(body);
				itemReport.FileResult = string.Empty;
			}

			if (itemReport.ShouldBeSavedToFolder)
				SaveReportToFolder();

			if (Logging.bw != null)
				if (MessageBox.Show("Отправить сообщение с отчетом следующим адресатам?" +
					Environment.NewLine + Environment.NewLine + itemReport.Recipients,
					"Отправка сообщения", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
					return;

			//if (Debugger.IsAttached)
			//	return;

			string recipients = itemReport.Recipients;
			if (hasError)
				recipients = Configuration.Instance.MailAdminAddress;

			Mail.SendMail(subject, body, recipients, itemReport.FileResult);
			Logging.ToLog("Завершение работы");

			return;
		}


		private static void WriteOutAcceptedParameters() {
			string message = Environment.NewLine + "Формат указания параметров:" + Environment.NewLine +
				"ID_отчета СмещениеДатаНачала СмещениеДатаОкончания (пример: 'FreeCells 0 6')" + Environment.NewLine +
				"ID_отчета ДатаНачала ДатаОкончания (пример: 'FreeCells 01.01.2018 31.01.2018')" +
				"ID_отчета PreviousMonth (пример: 'FreeCells PreviousMonth' - отчет за предыдущий месяц)" +
				Environment.NewLine + Environment.NewLine +
				"Варианты отчетов:" + Environment.NewLine;

			foreach (ItemReport item in Configuration.Instance.ReportItems)
				message += "ID: " + item.ID + " (" + item.Name + ")" + Environment.NewLine;

			Logging.ToLog(message);
		}

		private static void ParseDateInterval(string[] args) {
			DateTime? dateBegin = null;
			DateTime? dateEnd = null;

			if (args.Length == 2) {
				if (args[1].Equals("PreviousMonth")) {
					dateBegin = DateTime.Now.AddMonths(-1).AddDays(-1 * (DateTime.Now.Day - 1));
					dateEnd = dateBegin.Value.AddDays(
						DateTime.DaysInMonth(dateBegin.Value.Year, dateBegin.Value.Month) - 1);
				}
			} else if (args.Length == 3) {
				if (int.TryParse(args[1], out int dateBeginOffset) &&
					int.TryParse(args[2], out int dateEndOffset)) {
					dateBegin = DateTime.Now.AddDays(dateBeginOffset);
					dateEnd = DateTime.Now.AddDays(dateEndOffset);
				} else if (DateTime.TryParseExact(args[1], "dd.MM.yyyy", CultureInfo.InvariantCulture,
					DateTimeStyles.None, out DateTime dateBeginArg) &&
					DateTime.TryParseExact(args[2], "dd.MM.yyyy", CultureInfo.InvariantCulture,
					DateTimeStyles.None, out DateTime dateEndArg)) {
					dateBegin = dateBeginArg;
					dateEnd = dateEndArg;
				}
			} else
				return;

			if (dateBegin.HasValue && dateEnd.HasValue)
				itemReport.SetPeriod(dateBegin.Value, dateEnd.Value);
		}


		private static void LoadData(FirebirdClient firebirdClient) {
			parameters = new Dictionary<string, object>() {
				{ "@dateBegin", dateBeginStr },
				{ "@dateEnd", dateEndStr }
			};

			Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " + dateBeginStr + " по " + dateEndStr);

			try {
				dataTableMainData = firebirdClient.GetDataTable(itemReport.Query, parameters, true);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				hasError = true;
			}

			Logging.ToLog("Получено строк: " + dataTableMainData.Rows.Count);
		}

		private static void WriteDataToFile() {
			if (dataTableMainData.Rows.Count > 0) {
				Logging.ToLog("Запись данных в файл");

				if (itemReport.SaveAsCSV)
					itemReport.FileResult = ExcelGeneral.SaveAsCSV(dataTableMainData, subject);
				else 
					itemReport.FileResult = ExcelGeneral.WriteDataTableToExcel(
						dataTableMainData, subject);
				
				if (File.Exists(itemReport.FileResult)) {
					body = "Отчет во вложении";
					Logging.ToLog("Данные сохранены в файл: " + itemReport.FileResult);
				} else {
					body = "Не удалось записать данные в файл: " + itemReport.FileResult;
					hasError = true;
				}
			} else {
				body = "Отсутствуют данные за период " + 
					itemReport.DateBegin.ToShortDateString() + "-" + 
					itemReport.DateEnd.ToShortDateString();
				hasError = true;
			}
		}

		private static void SaveReportToFolder() {
			if (hasError)
				return;

			if (string.IsNullOrEmpty(itemReport.FolderToSave)) {
				Logging.ToLog("!!! Не указан путь сохранения, пропуск");
				return;
			}

			try {
				body = "Файл с отчетом сохранен по адресу: " + Environment.NewLine +
					SaveFileToNetworkFolder();

				itemReport.FileResult = string.Empty;
			} catch (Exception e) {
				Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
				body = "Не удалось сохранить отчет в папку " + itemReport.FolderToSave +
					Environment.NewLine + e.Message + Environment.NewLine + e.StackTrace;
			}
		}

		public static string SaveFileToNetworkFolder() {
			string fileName = Path.GetFileName(itemReport.FileResult);
			string destFile = Path.Combine(itemReport.FolderToSave, fileName);
			File.Copy(itemReport.FileResult, destFile, true);

			return "<a href=\"" + itemReport.FolderToSave + "\">" + itemReport.FolderToSave + "</a>";
		}
	}
}
