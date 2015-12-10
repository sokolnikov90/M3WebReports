using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml.Linq;

using M3Atms;
using M3Incidents;
using M3Dictionaries;

namespace M3Reports
{
    using M3IPClient;

    public class ReportIncidentsByTypes : ReportBuilder
    {
        private List<ReportColumns> reportColumns = new List<ReportColumns>();

        internal override void MakeAnExcel()
        {
            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\INC_TYPES_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(this.Info.path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart;
                WorksheetPart worksheetPart;
                WorkbookStylesPart workbookStylesPart;

                workbookpart = spreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                workbookStylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = M3Utils.ExcelHelper.MakeStyleSheet();

                Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                this.FillReportColumns();

                for (int i = 0; i < this.Data.DictionariesGet.Types.Count; i++)
                {
                    if (this.Data.DictionariesGet.Types[i].text == "ATMLocked")
                        continue;
                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheet sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                        Name = this.Data.DictionariesGet.Types[i].text,
                        SheetId = (uint)(i + 1)
                    };

                    sheets.Append(sheet);

                    this.CreateHeaderRow(worksheetPart, this.Info.from, this.Info.to);

                    this.CreateDataRows(worksheetPart, this.Data.DictionariesGet.Types[i]);

                    for (int j = 1; j <= this.reportColumns.Count(); j++)
                        M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, j, this.reportColumns[j - 1].width);
                }

                workbookpart.Workbook.Save();
            }
        }

        private void FillReportColumns()
        {
            string xmlPath = String.Empty;
            switch (M3UserSession.BankName)
            {
                case "RNCB":
                    xmlPath = @"bin/M3Reports/ReportIncidentColumn-RNCB.xml";
                    break;
                case "BM":
                    xmlPath = @"bin/M3Reports/ReportIncidentColumn-BM.xml";
                    break;
                default:
                    xmlPath = @"bin/M3Reports/ReportIncidentColumn.xml";
                    break;
            }
            this.reportColumns = ReportDataProvider.ParseXML(xmlPath);
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart, string from, string to)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = string.Empty;

            switch (this.Info.type)
            {
                case "IncidentsHistoryCurrent":
                    title = ReportsSource.ReportOnTheEliminationOfIncidents;
                    break;
                case "IncidentsHistoryRange":
                    title = string.Join(" ", new []
                            {
                                ReportsSource.ReportOnTheEliminationOfIncidents,
                                ReportsSource.From,
                                from,
                                ReportsSource.To,
                                to
                            });
                    break;
            }

            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                if (i > 1)
                    title = string.Empty;

                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportColumns[0].localtion + row.RowIndex, this.reportColumns[this.reportColumns.Count - 1].localtion + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
        }

        private void CreateDataRows(WorksheetPart worksheetPart, GetTypes.Data type)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            DateTime date;
            double hours = 0;
            row = (Row)sheetData.LastChild;
            string Status;
            string number;
            var incidents = this.Data.Incidents.Where(incident => incident.typeId == type.id.ToString());

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                Atm = this.Data.AtmInfo.First(atm => atm.Id == incident.atmId);
                number = incident.timeCreated.Substring(2, 8).Replace("-", "") + incident.id;
                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, number, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.Vizname, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Model, CellValues.String, 5U);
                if (M3UserSession.BankName == "BM")
                    M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, Atm.Institute, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, Atm.Place, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 7, row.RowIndex, type.text, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.timeCreated, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 9, row.RowIndex, incident.timeRegistrationService, CellValues.String, 5U);

                Status = this.Data.DictionariesGet.Statuses.First(inc => inc.id == Convert.ToInt32(incident.statusId)).text;

                date = DateTime.Parse(incident.timeCreated);
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                M3Utils.ExcelHelper.CreateCell(row, 10, row.RowIndex, date.ToString("yyyy-MM-dd hh:mm:ss"), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 11, row.RowIndex, Status, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 12, row.RowIndex, this.GetIncidentSubject(Convert.ToInt32(incident.id)), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 13, row.RowIndex, incident.comments, CellValues.String, 5U);
                if (M3UserSession.BankName == "RNCB")                
                    M3Utils.ExcelHelper.CreateCell(row, 14, row.RowIndex, ((incident.isCritical == 1) ? ReportsSource.Yes : ReportsSource.No), CellValues.String, 5U);
            }
            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }

        }

        private string GetIncidentSubject(int incidentId)
        {
            for (int i = 0; i < this.Data.DictionariesInfo.incidentsRules.data.Count; i++)
            {
                if (incidentId == this.Data.DictionariesInfo.incidentsRules.data[i].id)
                    return this.Data.DictionariesInfo.incidentsRules.data[i].iSubject;
            }

            return ReportsSource.Unknown;
        }
    }
}