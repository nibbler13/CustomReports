using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;

namespace CustomReports {
	public static class Logging {
		public static string AssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";
		private readonly static string LOG_FILE_NAME = Assembly.GetExecutingAssembly().GetName().Name + "_*.log";
		public static BackgroundWorker bw;

		public static void ToLog(string msg) {
			string today = DateTime.Now.ToString("yyyyMMdd");
			string logFileName = AssemblyDirectory + LOG_FILE_NAME.Replace("*", today);

			try {
				using (System.IO.StreamWriter sw = System.IO.File.AppendText(logFileName)) {
					string logLine = System.String.Format("{0:G}: {1}", System.DateTime.Now, msg);
					sw.WriteLine(logLine);

					if (bw != null)
						bw.ReportProgress(0, logLine);
				}
			} catch (Exception e) {
				Console.WriteLine("LogMessageToFile exception: " + logFileName + Environment.NewLine + e.Message + 
					Environment.NewLine + e.StackTrace);
			}

			Console.WriteLine(msg);
		}

		public static void CheckAndCleanOldFiles(uint maxLogfilesQuantity) {
			ToLog("Logging - Удаление старых лог-файлов");

			try {
				DirectoryInfo dirInfo = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
				FileInfo[] files = dirInfo.GetFiles(LOG_FILE_NAME).OrderBy(p => p.CreationTime).ToArray();

				if (files.Length <= maxLogfilesQuantity)
					return;

				for (int i = 0; i < files.Length - maxLogfilesQuantity; i++)
					files[i].Delete();
			} catch (Exception e) {
				ToLog("Logging - CheckAndCleanOldFiles exception: " + e.Message + Environment.NewLine + e.StackTrace);
			}
		}
	}
}
