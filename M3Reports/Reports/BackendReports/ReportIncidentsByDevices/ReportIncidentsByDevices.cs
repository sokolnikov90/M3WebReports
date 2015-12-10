namespace M3Reports
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using M3Atms;

    using M3Dictionaries;

    using M3Incidents;

    using M3IPClient;

    using M3Utils;

    public class ReportIncidentsByDevices : ReportBuilder
    {
        protected List<ReportColumns> reportColumns = new List<ReportColumns>();

        internal override void MakeAnExcel()
        {
            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\INC_DEVICES_"
                             + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_"
                             + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

            using (
                SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(
                    this.Info.path,
                    SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart;
                WorksheetPart worksheetPart;
                WorkbookStylesPart workbookStylesPart;

                workbookpart = spreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();



                workbookStylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = ExcelHelper.MakeStyleSheet();

                Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                this.reportColumns = ReportDataProvider.ParseXML(String.Format(@"bin/M3Reports/ReportIncidentColumn{0}.xml", M3UserSession.BankName));

                worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheet sheet = new Sheet()
                                  {
                                      Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                                      Name = ReportsSource.Incidents,
                                      SheetId = (uint)(1)
                                  };

                sheets.Append(sheet);

                this.FillReportColumns();

                this.CreateHeaderRow(worksheetPart, this.Info.from, this.Info.to);

                foreach (GetDevicesTypes.Data item in this.Data.DictionariesGet.DevicesTypes)
                {
                    this.CreateDataRows(worksheetPart, item);
                }

                for (int j = 1; j <= this.reportColumns.Count(); j++) ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, j, this.reportColumns[j - 1].width);
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
                if (i > 1) title = "";

                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }

            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportColumns[0].localtion + row.RowIndex, this.reportColumns[this.reportColumns.Count - 1].localtion + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
        }

        private void CreateDataRows(WorksheetPart worksheetPart, GetDevicesTypes.Data type)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            DateTime date;
            int hours = 0;
            row = (Row)sheetData.LastChild;
            string Status;
            string number;
            var incidents = this.Data.Incidents.Where(incident => Convert.ToInt32(incident.deviceTypeId) == type.id);
            if (!incidents.IsNullOrEmpty())
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 30D, CustomHeight = true });
                row = (Row)sheetData.LastChild;
                var title = string.Empty;
                for (int i = 1; i <= this.reportColumns.Count; i++)
                {
                    if (i > 1) title = string.Empty;
                    else title = type.description;
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }

                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportColumns[0].localtion + row.RowIndex, this.reportColumns[this.reportColumns.Count - 1].localtion + row.RowIndex);
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                for (int i = 1; i <= this.reportColumns.Count; i++)
                {
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
                }

                foreach (Incident incident in incidents)
                {
                    sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                    row = (Row)sheetData.LastChild;

                    Atm = this.Data.AtmInfo.First(atm => atm.Id == incident.atmId);
                    number = incident.timeCreated.Substring(2, 8).Replace("-", "") + incident.id;

                    int columnIndex = 1;

                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, number, CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, Atm.Vizname, CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, Atm.Model, CellValues.String, 5U);
                    if (M3UserSession.BankName == "BM")
                        ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, Atm.Institute, CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, Atm.Place, CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, type.description, CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, incident.timeCreated, CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, incident.timeRegistrationService, CellValues.String, 5U);

                    Status = this.Data.DictionariesGet.Statuses.First(inc => inc.id == Convert.ToInt32(incident.statusId)).text;

                    date = DateTime.Parse(incident.timeCreated);

                    if (Int32.TryParse(Atm.RecoveryTime, out hours)) date.AddHours(hours);

                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, date.ToString("yyyy-MM-dd hh:mm:ss"), CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, Status, CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data), CellValues.String, 5U);
                    ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, incident.comments, CellValues.String, 5U);
                    if (M3UserSession.BankName == "RNCB")
                        ExcelHelper.CreateCell(row, columnIndex++, row.RowIndex, ((incident.isCritical == 1) ? ReportsSource.Yes : ReportsSource.No), CellValues.String, 5U);
                }
            }
        }
    }
}
