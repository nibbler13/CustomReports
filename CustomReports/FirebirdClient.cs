using System;
using System.Data;
using System.Collections.Generic;
using FirebirdSql.Data.FirebirdClient;

namespace CustomReports {
    sealed class FirebirdClient : IDisposable {
		private readonly FbConnection connection;

		public FirebirdClient(string ipAddress, string baseName, string user, string pass, bool isGui = false) {
			Logging.ToLog("FirebirdClient - Создание подключения к базе: " + 
				ipAddress + ":" + baseName);

			FbConnectionStringBuilder cs = new FbConnectionStringBuilder();

			try {
				cs.DataSource = ipAddress;
				cs.Database = baseName;
				cs.UserID = user;
				cs.Password = pass;
				cs.Charset = "NONE";
				cs.Pooling = false;

				connection = new FbConnection(cs.ToString());
			} catch (Exception e) {
				Logging.ToLog("FirebirdClient - " + e.Message + Environment.NewLine + e.StackTrace);

				if (isGui)
					throw;
			}
		}



		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2100:Review SQL queries for security vulnerabilities")]
		public DataTable GetDataTable(string query, Dictionary<string, object> parameters = null, bool hasToRethrowException = false) {
			DataTable dataTable = new DataTable();

			try {
				connection.Open();
#pragma warning disable CA2100 // Review SQL queries for security vulnerabilities
				using (FbCommand command = new FbCommand(query, connection)) {
					if (parameters != null && parameters.Count > 0)
						foreach (KeyValuePair<string, object> parameter in parameters)
							command.Parameters.AddWithValue(parameter.Key, parameter.Value);

					using (FbDataAdapter fbDataAdapter = new FbDataAdapter(command))
						fbDataAdapter.Fill(dataTable);
				}
#pragma warning restore CA2100 // Review SQL queries for security vulnerabilities
			} catch (Exception e) {
				Logging.ToLog("FirebirdClient - GetDataTable exception: " + query + 
					Environment.NewLine + e.Message + Environment.NewLine + e.StackTrace);

				if (hasToRethrowException)
					throw;
			} finally {
				connection.Close();
			}

			return dataTable;
		}

		public void Dispose() {
			connection.Dispose();
		}
	}
}
