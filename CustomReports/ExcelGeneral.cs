using NPOI.SS.UserModel;
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

		public static string GetResultFilePath(string resultFilePrefix, bool isPlainText = false) {
			string resultPath = Path.Combine(Program.AssemblyDirectory, "Results");
			if (!Directory.Exists(resultPath))
				Directory.CreateDirectory(resultPath);

			foreach (char item in Path.GetInvalidFileNameChars())
				resultFilePrefix = resultFilePrefix.Replace(item, '-');

			string fileEnding = ".xlsx";
			if (isPlainText)
				fileEnding = ".txt";

			string resultFile = Path.Combine(resultPath, resultFilePrefix + " " + DateTime.Now.ToString("yyyyMMdd_HHmmss") + fileEnding);
			
			return resultFile;
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
			cellStyle.BorderTop = BorderStyle.Dotted;
			cellStyle.BorderBottom = BorderStyle.Dotted;
			cellStyle.BorderLeft = BorderStyle.Dotted;
			cellStyle.BorderRight = BorderStyle.Dotted;
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

					if (double.TryParse(value, out double result)) {
						cell.SetCellValue(result);
					} else if (value.Length >= 8 && DateTime.TryParse(value, out DateTime date)) {
						cell.SetCellValue(date);

						if (date.TimeOfDay.TotalSeconds > 0) {
							cell.CellStyle = cellStyleDateWithTime;
						} else {
							cell.CellStyle = cellStyleOnlyDate;
						}
					} else if (value.Length >=3 && value.Length <= 5 &&
						DateTime.TryParse(value, out DateTime time)) {
						cell.SetCellValue(time);
						cell.CellStyle = cellStyleOnlyTime;
					} else {
						cell.SetCellValue(value);
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
