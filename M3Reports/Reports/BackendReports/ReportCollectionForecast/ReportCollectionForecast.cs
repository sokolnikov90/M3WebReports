namespace M3Reports
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web;
    using System.Xml.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using M3Atms;
    using M3Incidents;

    using M3Utils;

    public class ReportCollectionForecast : ReportBuilder
    {
        public List<Info> atmInfo;

        public Dictionary<int, WithdrawalData> atmWithdraw;

        private List<ReportColumns> reportPage1Columns = new List<ReportColumns>();
        private List<ReportColumns> reportPageNColumns = new List<ReportColumns>();

        internal override void MakeAnExcel()
        {
            this.Info.path = this.Info.path.Replace("/", "\\") + "\\ATMS_FORECAST_"
                                     + this.Info.from.Replace("-", "").Substring(2) + "_"
                                     + this.Info.to.Replace("-", "").Substring(2) + ".xlsx";

            using (
                SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(this.Info.path,
                    SpreadsheetDocumentType.Workbook)) // path, SpreadsheetDocumentType.Workbook, false, openSettings))
            {
                WorkbookPart workbookpart;
                WorksheetPart worksheetPart;
                WorkbookStylesPart workbookStylesPart;

                workbookpart = spreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                workbookStylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = ExcelHelper.MakeStyleSheet();

                this.reportPage1Columns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportCollectionForecastPage1Column.xml");
                this.reportPageNColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportCollectionForecastPageNColumn.xml");

                Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                uint uId = 1;

                #region Sheet#1

                worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheet sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    Name = ReportsSource.GeneralForecast,
                    SheetId = uId
                };

                sheets.Append(sheet);

                this.CreatePage1HeaderRow(worksheetPart);
                this.CreatePage1DataRows(worksheetPart, this.atmInfo, this.atmWithdraw);

                for (int i = 1; i <= this.reportPage1Columns.Count(); i++)
                    ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportPage1Columns[i - 1].width);

                #endregion

                #region Sheet#N

                uId++;

                foreach (var oneAtmWithdraw in this.atmWithdraw)
                {
                    var oneAtmInfo = this.atmInfo.First(a => a.Id == oneAtmWithdraw.Key.ToString());

                    if (oneAtmInfo == null) continue;

                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    sheet = new Sheet()
                                {
                                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                                    Name = oneAtmInfo.DeviceNumber,
                                    SheetId = uId
                                };

                    sheets.Append(sheet);

                    uId++;

                    this.CreatePageNHeaderRows(worksheetPart, oneAtmInfo, this.reportPageNColumns);

                    this.CreatePageNDataRows(worksheetPart, oneAtmWithdraw, oneAtmInfo);

                    for (int i = 1; i <= this.reportPageNColumns.Count(); i++) ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportPageNColumns[i - 1].width);
                }
                #endregion

                workbookpart.Workbook.Save();
            }
        }

        private void CreatePage1HeaderRow(WorksheetPart worksheetPart)
        {
            try
            {
                Row row;
                SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

                sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                string title = string.Join(" ", new[]
                        {
                            ReportsSource.WithdrawCashForecast,
                            ReportsSource.From, this.Info.from,
                            ReportsSource.To, this.Info.to
                        });
                    
                //"Прогноз снятия наличных на банкоматах с " + this.Info.from + " по " + this.Info.to;

                for (int i = 1; i <= this.reportPage1Columns.Count; i++)
                {
                    if (i > 1)
                        title = "";

                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }

                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportPage1Columns[0].localtion + row.RowIndex, this.reportPage1Columns[this.reportPage1Columns.Count - 1].localtion + row.RowIndex);

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 30D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                for (int i = 1; i <= this.reportPage1Columns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportPage1Columns[i - 1].title, CellValues.String, 4U);
            }
            catch (Exception exp)
            {
                Log.Instance.Info(this + ".CreatePage1HeaderRow() exeption: " + exp.Message);
            }
        }

        private void CreatePage1DataRows(WorksheetPart worksheetPart, List<Info> atmInfo, Dictionary<int, WithdrawalData> atmWithdrawal)
        {
            try
            {
                Row row;

                SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

                row = (Row)sheetData.LastChild;

                foreach (var oneAtmWithdraw in this.atmWithdraw)
                {
                    var oneAtmInfo = this.atmInfo.First(a => a.Id == oneAtmWithdraw.Key.ToString());

                    if (oneAtmInfo == null) continue;

                    int lastRemain = 0;

                    string firstTimeMinAmount = string.Empty;
                    string firstTimeZeroAmount = string.Empty;

                    foreach (var dateTimeWithdraw in oneAtmWithdraw.Value.Data)
                    {
                        //if ((dateTimeWithdraw.Key < from) || (dateTimeWithdraw.Key > to))
                        //    continue;

                        if (dateTimeWithdraw.Value[1] != -1)
                        {
                            lastRemain = dateTimeWithdraw.Value[1];
                        }
                        else
                        {
                            lastRemain -= dateTimeWithdraw.Value[2];

                            if (lastRemain > 0 && lastRemain <= Convert.ToInt32(oneAtmInfo.MinAmount))
                            {
                                if (firstTimeMinAmount == string.Empty) firstTimeMinAmount = dateTimeWithdraw.Key.ToString("yyyy-MM-dd");
                            }

                            if (lastRemain <= 0)
                            {
                                if (firstTimeZeroAmount == string.Empty) firstTimeZeroAmount = dateTimeWithdraw.Key.ToString("yyyy-MM-dd");
                            }
                        }
                    }

                    sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 30D, CustomHeight = true });
                    row = (Row)sheetData.LastChild;

                    ExcelHelper.CreateCell(row, 1, row.RowIndex, oneAtmInfo.Vizname, CellValues.String, 2U);
                    ExcelHelper.CreateCell(row, 2, row.RowIndex, oneAtmInfo.MinAmount, CellValues.String, 2U);

                    if ((firstTimeMinAmount == string.Empty) && (firstTimeZeroAmount == string.Empty))
                    {
                        ExcelHelper.CreateCell(row, 3, row.RowIndex, "Не ранее " + this.Info.to, CellValues.String, 8U);
                        ExcelHelper.CreateCell(row, 4, row.RowIndex, "Не ранее " + this.Info.to, CellValues.String, 9U);
                    }
                    else if ((firstTimeMinAmount != string.Empty) && (firstTimeZeroAmount == string.Empty))
                    {
                        ExcelHelper.CreateCell(row, 3, row.RowIndex, firstTimeMinAmount, CellValues.String, 8U);
                        ExcelHelper.CreateCell(row, 4, row.RowIndex, "Не ранее " + this.Info.to, CellValues.String, 9U);                        
                    }
                    else if ((firstTimeMinAmount == string.Empty) && (firstTimeZeroAmount != string.Empty))
                    {
                        ExcelHelper.CreateCell(row, 3, row.RowIndex, firstTimeZeroAmount, CellValues.String, 8U);
                        ExcelHelper.CreateCell(row, 4, row.RowIndex, firstTimeZeroAmount, CellValues.String, 9U);
                    }
                    else // All not empty
                    {
                        ExcelHelper.CreateCell(row, 3, row.RowIndex, firstTimeMinAmount, CellValues.String, 8U);
                        ExcelHelper.CreateCell(row, 4, row.RowIndex, firstTimeZeroAmount, CellValues.String, 9U);                        
                    }
                }
            }
            catch (Exception exp)
            {
                Log.Instance.Info(this + ".CreatePage1DataRows() exeption: " + exp.Message);
            }
        }

        private void CreatePageNHeaderRows(WorksheetPart worksheetPart, Info atmInfo, List<ReportColumns> reportColumns)
        {
            try
            {
                Row row;
                SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

                sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                string title = "История и прогноз снятия наличных на банкомате : " + Environment.NewLine + atmInfo.Vizname;

                for (int i = 1; i <= reportColumns.Count; i++)
                {
                    if (i > 1)
                        title = "";

                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }

                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, reportColumns[0].localtion + row.RowIndex, reportColumns[reportColumns.Count - 1].localtion + row.RowIndex);

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 30D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                for (int i = 1; i <= reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, reportColumns[i - 1].title, CellValues.String, 4U);
            }
            catch (Exception exp)
            {
                Log.Instance.Info(
                    this + ".CreatePageNDataRows(...) exception:",
                    exp.Message,
                    exp.Source,
                    exp.StackTrace);
            }
        }

        private void CreatePageNDataRows(WorksheetPart worksheetPart, KeyValuePair<int, WithdrawalData> data, Info atmInfo)
        {
            try
            {
                Row row;

                SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

                row = (Row)sheetData.LastChild;

                int lastRemain = 0;

                foreach (var dateTimeWithdraw in data.Value.Data)
                {
                    //if ((dateTimeWithdraw.Key < from) || (dateTimeWithdraw.Key > to))
                    //    continue;

                    sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
                    row = (Row)sheetData.LastChild;

                    ExcelHelper.CreateCell(row, 1, row.RowIndex, dateTimeWithdraw.Key.ToString("yyyy-MM-dd"), CellValues.String, 2U);
                    ExcelHelper.CreateCell(row, 2, row.RowIndex, dateTimeWithdraw.Key.DayOfWeekRus().ToString(), CellValues.String, 2U);
                    ExcelHelper.CreateCell(row, 3, row.RowIndex, dateTimeWithdraw.Value[0] != -1 ? dateTimeWithdraw.Value[0].ToString() : "-", CellValues.String, 2U);
                    ExcelHelper.CreateCell(row, 4, row.RowIndex, dateTimeWithdraw.Value[2] != -1 ? dateTimeWithdraw.Value[2].ToString() : "-", CellValues.String, 2U);
                    ExcelHelper.CreateCell(row, 5, row.RowIndex, dateTimeWithdraw.Value[1] != -1 ? dateTimeWithdraw.Value[1].ToString() : "-", CellValues.String, 2U);

                    if (dateTimeWithdraw.Value[1] != -1)
                    {
                        ExcelHelper.CreateCell(row, 6, row.RowIndex, "-", CellValues.String, 2U);
                        lastRemain = dateTimeWithdraw.Value[1];
                    }
                    else
                    {
                        lastRemain -= dateTimeWithdraw.Value[2];

                        UInt32 styleIndex;

                        if (lastRemain > Convert.ToInt32(atmInfo.MinAmount))
                            styleIndex = 7U;
                        else if (lastRemain > 0 && lastRemain <= Convert.ToInt32(atmInfo.MinAmount))
                            styleIndex = 8U;
                        else
                            styleIndex = 9U;

                        ExcelHelper.CreateCell(row, 6, row.RowIndex, lastRemain < 0 ? "0" : lastRemain.ToString(), CellValues.String, styleIndex);
                    }
                }
            }
            catch (Exception exp)
            {
                Log.Instance.Info(
                    this + ".CreatePageNDataRows(...) exception:",
                    exp.Message,
                    exp.Source,
                    exp.StackTrace);
            }
        }
    }
}