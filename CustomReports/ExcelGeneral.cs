﻿using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows;

namespace CustomReports {
	class ExcelGeneral {
		//============================ NPOI Excel ============================
		protected static bool CreateNewIWorkbook(string resultFilePrefix, out IWorkbook workbook, out ISheet sheet, out string resultFile) {
			workbook = null;
			sheet = null;
			resultFile = string.Empty;

			try {
				string resultPath = GetResultFilePath(resultFilePrefix);
				workbook = new XSSFWorkbook();
				sheet = workbook.CreateSheet("Данные");
				resultFile = resultPath;
				return true;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}

		public static string GetResultFilePath(string resultFilePrefix, bool isCSV = false, bool isXML = false) {
			string resultPath = Path.Combine(Program.AssemblyDirectory, "Results");
			if (!Directory.Exists(resultPath))
				Directory.CreateDirectory(resultPath);

			foreach (char item in Path.GetInvalidFileNameChars())
				resultFilePrefix = resultFilePrefix.Replace(item, '-');

			string fileEnding = ".xlsx";
			if (isCSV)
				fileEnding = ".csv";
			else if (isXML)
				fileEnding = ".xml";

			string resultFile = Path.Combine(resultPath, resultFilePrefix + " " + DateTime.Now.ToString("yyyyMMdd_HHmmss") + fileEnding);
			
			return resultFile;
		}

		public static string SaveAsCSV(DataTable dataTable, string resultFilePrefix) {
			string fileName = GetResultFilePath(resultFilePrefix, isCSV: true);

			StringBuilder sb = new StringBuilder();
			IEnumerable<string> columnNames = dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName);
			sb.AppendLine(string.Join(";", columnNames));
			foreach (DataRow dataRow in dataTable.Rows) {
				IEnumerable<string> fields = dataRow.ItemArray.Select(f => f.ToString());
				sb.AppendLine(string.Join(";", fields));
			}

			try {
				File.WriteAllText(fileName, sb.ToString(), Encoding.GetEncoding("windows-1251"));
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				return string.Empty;
			}

			return fileName;
		}

		public static string SaveAsXML(DataTable dataTable, string resultFilePrefix) {
			string fileName = GetResultFilePath(resultFilePrefix, isXML: true);

			StringBuilder sb = new StringBuilder();

			sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
			sb.AppendLine("<?xml-stylesheet type=\"text/xsl\" href=\"" + GetHtmlString(fileName) + "\"?>");
			sb.AppendLine("<CACHE>");
			sb.AppendLine("<TITLE>" + resultFilePrefix + "</TITLE>");
			sb.AppendLine("<STYLES>");
			sb.AppendLine("<STYLE Id=\"0\" AlignText=\"Center\" FontName=\"Tahoma\" FontCharset=\"1\" Bold=\"False\" Italic=\"False\" Underline=\"False\" StrikeOut=\"False\" FontColor=\"rgb(0,0,0)\" FontSize=\"8\" BrushStyle=\"Solid\" BrushBkColor=\"rgb(240,240,240)\" BrushFgColor=\"rgb(0,0,0)\">");
			sb.AppendLine("<BORDER_LEFT IsDefault=\"False\" Color=\"rgb(160,160,160)\" Width=\"1\"/>");
			sb.AppendLine("<BORDER_UP IsDefault=\"False\" Color=\"rgb(160,160,160)\" Width=\"1\"/>");
			sb.AppendLine("<BORDER_RIGHT IsDefault=\"False\" Color=\"rgb(160,160,160)\" Width=\"1\"/>");
			sb.AppendLine("<BORDER_DOWN IsDefault=\"False\" Color=\"rgb(160,160,160)\" Width=\"1\"/>");
			sb.AppendLine("</STYLE>");
			sb.AppendLine("<STYLE Id=\"1\" AlignText=\"Center\" FontName=\"Tahoma\" FontCharset=\"1\" Bold=\"False\" Italic=\"False\" Underline=\"False\" StrikeOut=\"False\" FontColor=\"rgb(0,0,0)\" FontSize=\"8\" BrushStyle=\"Solid\" BrushBkColor=\"rgb(255,255,255)\" BrushFgColor=\"rgb(0,0,0)\">");
			sb.AppendLine("<BORDER_LEFT IsDefault=\"False\" Color=\"rgb(192,192,192)\" Width=\"1\"/>");
			sb.AppendLine("<BORDER_UP IsDefault=\"False\" Color=\"rgb(192,192,192)\" Width=\"1\"/>");
			sb.AppendLine("<BORDER_RIGHT IsDefault=\"False\" Color=\"rgb(192,192,192)\" Width=\"1\"/>");
			sb.AppendLine("<BORDER_DOWN IsDefault=\"False\" Color=\"rgb(192,192,192)\" Width=\"1\"/>");
			sb.AppendLine("</STYLE>");
			sb.AppendLine("<STYLE Id=\"2\" AlignText=\"Left\" FontName=\"Tahoma\" FontCharset=\"1\" Bold=\"False\" Italic=\"False\" Underline=\"False\" StrikeOut=\"False\" FontColor=\"rgb(0,0,0)\" FontSize=\"8\" BrushStyle=\"Solid\" BrushBkColor=\"rgb(255,255,255)\" BrushFgColor=\"rgb(0,0,0)\">");
			sb.AppendLine("<BORDER_LEFT IsDefault=\"False\" Color=\"rgb(192,192,192)\" Width=\"1\"/>");
			sb.AppendLine("<BORDER_UP IsDefault=\"False\" Color=\"rgb(192,192,192)\" Width=\"1\"/>");
			sb.AppendLine("<BORDER_RIGHT IsDefault=\"False\" Color=\"rgb(192,192,192)\" Width=\"1\"/>");
			sb.AppendLine("<BORDER_DOWN IsDefault=\"False\" Color=\"rgb(192,192,192)\" Width=\"1\"/>");
			sb.AppendLine("</STYLE>");
			sb.AppendLine("</STYLES>");
			sb.AppendLine("<LINES ColCount=\"" + dataTable.Columns.Count + "\" RowCount=\"" + (dataTable.Rows.Count + 1) + "\">");

			sb.AppendLine("<LINE Height=\"19\">");
			foreach (DataColumn column in dataTable.Columns)
				sb.AppendLine("<CELL StyleClass=\"0\" Width=\"70\" Align=\"Center\">" + GetHtmlString(column.ColumnName) + "</CELL>");

			sb.AppendLine("</LINE>");
			sb.AppendLine("");

			foreach (DataRow row in dataTable.Rows) {
				sb.AppendLine("<LINE Height=\"18\">");
				foreach (object item in row.ItemArray)
					sb.AppendLine("<CELL StyleClass=\"1\" Width=\"70\" Align=\"Center\">" + GetHtmlString(item.ToString()) + "</CELL>");

				sb.AppendLine("</LINE>");
				sb.AppendLine("");
			}

			sb.AppendLine("</LINES>");
			sb.AppendLine("</CACHE>");

			try {
				File.WriteAllText(fileName, sb.ToString());
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				return string.Empty;
			}

			return fileName;
		}

		private static string GetHtmlString(string text) {
			return HttpUtility.HtmlEncode(text);
		}

		protected static bool SaveAndCloseIWorkbook(IWorkbook workbook, string resultFile) {
			try {
				using (FileStream stream = new FileStream(resultFile, FileMode.Create, FileAccess.Write))
					workbook.Write(stream);

				workbook.Close();

				return true;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}



		public static string WriteDataTableToExcel(DataTable dataTable, string resultFilePrefix) {
			IWorkbook workbook = null;
			ISheet sheet = null;
			string resultFile = string.Empty;

			if (!CreateNewIWorkbook(resultFilePrefix, out workbook, out sheet, out resultFile))
				return string.Empty;

			IFont fontMain = workbook.CreateFont();
			fontMain.FontName = "Calibri";
			fontMain.FontHeightInPoints = 10;
			IDataFormat dataFormat = workbook.CreateDataFormat();
			ICellStyle cellStyle = workbook.CreateCellStyle();
			cellStyle.BorderTop = BorderStyle.Thin;
			cellStyle.BorderBottom = BorderStyle.Thin;
			cellStyle.BorderLeft = BorderStyle.Thin;
			cellStyle.BorderRight = BorderStyle.Thin;
			cellStyle.SetFont(fontMain);

			ICellStyle cellStyleDateWithTime = workbook.CreateCellStyle();
			cellStyleDateWithTime.CloneStyleFrom(cellStyle);
			cellStyleDateWithTime.DataFormat = dataFormat.GetFormat("dd.MM.yy HH:mm");

			ICellStyle cellStyleOnlyDate = workbook.CreateCellStyle();
			cellStyleOnlyDate.CloneStyleFrom(cellStyle);
			cellStyleOnlyDate.DataFormat = dataFormat.GetFormat("dd.MM.yyyy");

			ICellStyle cellStyleOnlyTime = workbook.CreateCellStyle();
			cellStyleOnlyTime.CloneStyleFrom(cellStyle);
			cellStyleOnlyTime.DataFormat = dataFormat.GetFormat("HH:mm");

			ICellStyle cellStyleOnlyTimeWithSeconds = workbook.CreateCellStyle();
			cellStyleOnlyTimeWithSeconds.CloneStyleFrom(cellStyle);
			cellStyleOnlyTimeWithSeconds.DataFormat = dataFormat.GetFormat("HH:mm:ss");

			IFont fontBold = workbook.CreateFont();
			fontBold.Boldweight = (short)FontBoldWeight.Bold;
			fontBold.FontName = "Calibri";
			fontBold.FontHeightInPoints = 10;
			ICellStyle cellStyleHeader = workbook.CreateCellStyle();
			cellStyleHeader.CloneStyleFrom(cellStyle);
			cellStyleHeader.SetFont(fontBold);
			cellStyleHeader.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
			cellStyleHeader.FillPattern = FillPattern.SolidForeground;

			IRow rowHeader = null;
			try { rowHeader = sheet.GetRow(0); } catch (Exception) { }

			if (rowHeader == null)
				rowHeader = sheet.CreateRow(0);

			int columnNumber = 0;
			foreach (DataColumn columnHeader in dataTable.Columns) {
				ICell cell = null;
				try { cell = rowHeader.GetCell(columnNumber); } catch (Exception) { }

				if (cell == null)
					cell = rowHeader.CreateCell(columnNumber);

				cell.CellStyle = cellStyleHeader;
				cell.SetCellValue(columnHeader.ColumnName);
				columnNumber++;
			}

			int rowNumber = 1;
			columnNumber = 0;
			foreach (DataRow dataRow in dataTable.Rows) {
				IRow row = null;
				try { row = sheet.GetRow(rowNumber); } catch (Exception) { }

				if (row == null)
					row = sheet.CreateRow(rowNumber);
				
				foreach (DataColumn column in dataTable.Columns) {
					ICell cell = null;
					try { cell = row.GetCell(columnNumber); } catch (Exception) { }

					if (cell == null)
						cell = row.CreateCell(columnNumber);

					cell.CellStyle = cellStyle;

					string value = dataRow[column].ToString();

					Type columnType = column.DataType;
					if (columnType == typeof(DateTime)) {
						if (DateTime.TryParse(value, out DateTime dateTime)) {
							cell.SetCellValue(dateTime);
							if (dateTime.TimeOfDay.TotalSeconds > 0)
								cell.CellStyle = cellStyleDateWithTime;
							else
								cell.CellStyle = cellStyleOnlyDate;
						} else
							cell.SetCellValue(value);
					} else if (columnType == typeof(TimeSpan)) {
						if (TimeSpan.TryParse(value, out TimeSpan timeSpan)) {
							cell.SetCellValue(DateTime.Parse(value));
							cell.CellStyle = cellStyleOnlyTimeWithSeconds;
						} else
							cell.SetCellValue(value);
					} else if (columnType == typeof(double) ||
						columnType == typeof(float) ||
						columnType == typeof(long) || 
						columnType == typeof(int) ||
						columnType == typeof(short)) {
						if (double.TryParse(value, out double parsedDouble))
							cell.SetCellValue(parsedDouble);
						else
							cell.SetCellValue(value);
					} else {
						if (double.TryParse(value, out double result)) {
							cell.SetCellValue(result);
						} else if (
							(value.Length == 7 || value.Length == 8) &&
							value.Count(x => x.Equals(':')) == 2 &&
							DateTime.TryParse(value, out DateTime timeWithSecond)) {
							cell.SetCellValue(timeWithSecond);
							cell.CellStyle = cellStyleOnlyTimeWithSeconds;
						} else if (value.Length >= 8 && DateTime.TryParse(value, out DateTime date)) {
							cell.SetCellValue(date);

							if (date.TimeOfDay.TotalSeconds > 0) {
								cell.CellStyle = cellStyleDateWithTime;
							} else {
								cell.CellStyle = cellStyleOnlyDate;
							}
						} else if (value.Length >= 3 && value.Length <= 5 &&
							DateTime.TryParse(value, out DateTime time)) {
							cell.SetCellValue(time);
							cell.CellStyle = cellStyleOnlyTime;
						} else {
							cell.SetCellValue(value);
						}
					}
					
					columnNumber++;
				}

				columnNumber = 0;
				rowNumber++;
			}

			for (int i = 0; i < dataTable.Columns.Count; i++)
				sheet.AutoSizeColumn(i);

			if (!SaveAndCloseIWorkbook(workbook, resultFile))
				return string.Empty;

			return resultFile;
		}
    }
}
