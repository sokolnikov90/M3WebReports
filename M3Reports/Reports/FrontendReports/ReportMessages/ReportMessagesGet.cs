using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Web;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace M3Reports
{
    public class ReportMessagesGet : ReportBuilder
    {
        List<ReportColumns> reportColumns { get; set; }

        internal override void MakeAnExcel()
        {
            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\MSG_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(this.Info.path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart;
                WorksheetPart worksheetPart;
                WorkbookStylesPart workbookStylesPart;

                workbookpart = spreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                workbookStylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = M3Utils.ExcelHelper.MakeStyleSheet();

                Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                Sheet sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    Name = ReportsSource.MessageHistory,
                    SheetId = 1
                };

                sheets.Append(sheet);

                this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportMessageColums.xml");

                this.CreateHeaderRow(worksheetPart, this.Info.from, this.Info.to);


                this.CreateDataRows(worksheetPart);

                for (int i = 1; i <= this.reportColumns.Count(); i++)
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportColumns[i - 1].width);

                worksheetPart.Worksheet.Save();
            }
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart, string from, string to)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.AppendChild(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = String.Join(" ", new []
                                                {
                                                    ReportsSource.ReportMessages,
                                                    ReportsSource.From,
                                                    from,
                                                    ReportsSource.To,
                                                    to
                                                });

            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                if (i > 1)
                    title = "";

                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportColumns[0].localtion + row.RowIndex, this.reportColumns[this.reportColumns.Count - 1].localtion + row.RowIndex);

            sheetData.AppendChild(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 1U);
            }
        }

        private void CreateDataRows(WorksheetPart worksheetPart)
        {
            for (int n = 0; n < this.Data.AtmInfo.Count; n++)
            {
                Row row;
                SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

                row = (Row)sheetData.LastChild;
                sheetData.AppendChild(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                string title = ReportsSource.ATM + ": " + this.Data.AtmInfo[n].Vizname + "; " + ReportsSource.Address + ": " + this.Data.AtmInfo[n].GeoAddress;

                for (int j = 1; j <= this.reportColumns.Count; j++)
                {
                    if (j > 1) title = "";

                    M3Utils.ExcelHelper.CreateCell(row, j, row.RowIndex, title, CellValues.String, 1U);
                }

                M3Utils.ExcelHelper.MergeCellsInRange(
                    worksheetPart.Worksheet,
                    this.reportColumns[0].localtion + row.RowIndex,
                    this.reportColumns[this.reportColumns.Count - 1].localtion + row.RowIndex);

                for (int i = 0; i < this.Data.MessageHistoryGet.info.data.Count; i++)
                {
                    sheetData.AppendChild(new Row() { RowIndex = (row.RowIndex + 1) });
                    row = (Row)sheetData.LastChild;

                    foreach (var p in this.Data.MessageHistoryGet.info.data[i].GetType().GetProperties())
                    {
                        var index = this.reportColumns.FindIndex(x => x.name == p.Name);

                        if (index >= 0)
                        {
                            DateTime dateTime;

                            int direction;
                            string value;

                            switch (p.Name)
                            {
                                case "dateTime":
                                    dateTime = Convert.ToDateTime(p.GetValue(this.Data.MessageHistoryGet.info.data[i], null));

                                    M3Utils.ExcelHelper.CreateCell(row,(index + 1),row.RowIndex,Convert.ToString(dateTime),CellValues.String,2U);
                                    break;
                                case "direction":
                                    direction = Convert.ToInt32(p.GetValue(this.Data.MessageHistoryGet.info.data[i], null));

                                    value = ReportsSource.NotDetermined;

                                    if (direction == 0) value = ReportsSource.NotDetermined;
                                    else if (direction == 1) value = ReportsSource.fromAtm;
                                    else if (direction == 2) value = ReportsSource.fromHost;

                                    M3Utils.ExcelHelper.CreateCell(row,(index + 1),row.RowIndex,value,CellValues.String,2U);
                                    break;
                                default:
                                    value = Convert.ToString(p.GetValue(this.Data.MessageHistoryGet.info.data[i], null));

                                    M3Utils.ExcelHelper.CreateCell(row,(index + 1),row.RowIndex,value,CellValues.String,0U);
                                    break;
                            }
                        }
                    }

                    worksheetPart.Worksheet.Save();
                }
            }
        }
    }
}