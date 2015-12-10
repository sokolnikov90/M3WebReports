using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Linq;
using System.Xml;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using M3Atms;
using M3Incidents;
using M3Dictionaries;

namespace M3Reports
{
    using M3Utils;

    public class ReportMonitoring : ReportBuilder
    {
        private Dictionary<string, int> failuresDict;

        private List<ReportColumns> reportColumns;

        private void CreateFailuresDict()
        {
            this.failuresDict = new Dictionary<string, int>();

            foreach (var data in this.Data.DictionariesGet.DevicesTypes)
            {
                this.failuresDict.Add(data.name, data.id);
            }
        }

        internal override void MakeAnExcel()
        {
            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\MNT_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

            this.CreateFailuresDict();

            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(this.Info.path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart;
                WorksheetPart worksheetPart;
                WorkbookStylesPart workbookStylesPart;

                workbookpart = spreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                workbookStylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = ExcelHelper.MakeStyleSheet();
                Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheet sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    Name = ReportsSource.MonitoringService,
                    SheetId = (uint)1
                };
                sheets.Append(sheet);

                this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportMonitoringColumn.xml");

                this.CreateHeaderRow(worksheetPart);
                this.CreateCountsRows(worksheetPart);
                this.CreateCommRows(worksheetPart);
                this.CreateNotAvailableRows(worksheetPart);
                this.CreatePrinterRows(worksheetPart);
                this.CreateSLMRows(worksheetPart);
                this.CreateRPaperRows(worksheetPart);
                this.CreateJPaperRows(worksheetPart);


                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 1, 12);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 2, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 3, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 4, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 5, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 6, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 7, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 8, 34);
                workbookpart.Workbook.Save();
            }
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = String.Join(" ", new[]
                                        {
                                            ReportsSource.ReportMonitoringServiceATMs,
                                            ReportsSource.From,
                                            this.Info.from,
                                            ReportsSource.To,
                                            this.Info.to
                                        });
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A1", "H1");
        }

        private void CreateCountsRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            string title = ReportsSource.SimplyBecauseOfTheLackOfCash;
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;

            this.Data.AtmCounts = (from count in this.Data.AtmCounts
                                   where count.total_RemainCass_Pos1 == "0" && count.total_RemainCass_Pos2 == "0" && count.total_RemainCass_Pos3 == "0" && count.total_RemainCass_Pos4 == "0"
                                   select count).ToList();

            if (this.Data.AtmCounts.Count > 0)
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {
                title = ReportsSource.No;
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                for (int i = 1; i <= 8; i++)
                {
                    ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
                }
                return;
            }

            foreach (CountsGet.AtmCountsData count in this.Data.AtmCounts)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.First(atm => atm.Id == count.atmId);

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, ReportsSource.LackOfCash, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.Model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 7, row.RowIndex, "", CellValues.String, 5U);
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= 8; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }
        }

        private void CreateCommRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = ReportsSource.LackOfCommunication;
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            var incidents = this.Data.Incidents.Where(incident => Convert.ToInt32(incident.deviceTypeId) == this.failuresDict["Communications"]).OrderBy(inc=>inc.timeCreated);

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {
                title = ReportsSource.No;
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                for (int i = 1; i <= 8; i++)
                {
                    ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
                }
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.First(atm => atm.Id == incident.atmId);

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, ReportsSource.NoConnection, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.Model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);

            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= 8; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }

        }

        private void CreateNotAvailableRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            string title = ReportsSource.BreakingOtherReasons;
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            var statuses = new List<int>() { 4, 7, 8, 9, 10, 11, 12, 13, 14, 15 };

            var incidents = this.Data.Incidents.Where(incident => statuses.Contains(Convert.ToInt32(incident.statusId))).OrderBy(inc => inc.timeCreated); 

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {
                title = ReportsSource.No;
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                for (int i = 1; i <= 8; i++)
                {
                    ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
                }
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.First(atm => atm.Id == incident.atmId);

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.Model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= 8; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }
        }

        private void CreatePrinterRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = ReportsSource.FaultReceiptPrinter;
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            var incidents = this.Data.Incidents.Where(incident => Convert.ToInt32(incident.deviceTypeId) == this.failuresDict["ReceiptPrinter"] && !(incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data) == "Ч.Принтер: Бумага закончилась -> FLM ч.принтер" || incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data) == "Ч.Принтер: Мало бумаги -> FLM ч.принтер")).OrderBy(inc => inc.timeCreated);

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {

                title = ReportsSource.No;
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                for (int i = 1; i <= 8; i++)
                {
                    ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
                }
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.First(atm => atm.Id == incident.atmId);

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.Model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= 8; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }

        }

        private void CreateSLMRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = ReportsSource.WaitingForRepairs;
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            var type = this.Data.DictionariesGet.Types.Where(typ => typ.text == "SLM").First();
            var incidents = this.Data.Incidents.Where(incident => incident.typeId == type.id.ToString()).OrderBy(inc=>inc.timeCreated);

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {

                title = ReportsSource.No;
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                for (int i = 1; i <= 8; i++)
                {
                    ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
                }
                return;
            }



            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.First(atm => atm.Id == incident.atmId);

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.Model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);
                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= 8; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }

        }

        private void CreateRPaperRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = ReportsSource.ThePaperEnds_ReceiptPrinter;
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            var incidents = this.Data.Incidents.Where(incident => Convert.ToInt32(incident.deviceTypeId) == this.failuresDict["ReceiptPrinter"] && (incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data) == "Ч.Принтер: Бумага закончилась -> FLM ч.принтер" || incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data) == "Ч.Принтер: Мало бумаги -> FLM ч.принтер")).OrderBy(inc => inc.timeCreated); 

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {
                title = ReportsSource.No;
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                for (int i = 1; i <= 8; i++)
                {
                    ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
                }
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.First(atm => atm.Id == incident.atmId);

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.Model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= 8; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }

        }

        private void CreateJPaperRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = ReportsSource.ThePaperEnds_JournalPrinter;
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            var incidents = this.Data.Incidents.Where(incident => Convert.ToInt32(incident.deviceTypeId) == this.failuresDict["JournalPrinter"] && (incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data) == "Ж.Принтер: Мало бумаги -> FLM ж.принтер" || incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data) == "Ж.Принтер: Бумага закончилась -> FLM ж.принтер")).OrderBy(inc => inc.timeCreated);

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {

                title = ReportsSource.No;
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                for (int i = 1; i <= 8; i++)
                {
                    ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
                }
                return;
            }

          

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.First(atm => atm.Id == incident.atmId);

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.Model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss "), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= 8; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }
        }
    }
}