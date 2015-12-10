using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

using M3Atms;
using M3Incidents;
using M3Dictionaries;

namespace M3Reports
{
    using M3Utils;

    public class ReportOutOfService : ReportBuilder
    {
        private Dictionary<string, string> failuresDict;

        public Dictionary<string, GroupsGet.AtmGroup> groups;

        List<ReportColumns> reportColumns = new List<ReportColumns>();

        private void CreateFailuresDict()
        {
            this.failuresDict = new Dictionary<string, string>();
            foreach (GetDevicesTypes.Data data in this.Data.DictionariesInfo.devicesTypes.data)
            {
                this.failuresDict.Add(data.name, data.id.ToString());
            }
        }

        private void SearchGroupNames()
        {
            this.groups = new Dictionary<string, GroupsGet.AtmGroup>();
            foreach (string id in this.Info.atmsGroupsId)
            {
                this.groups.Add(id, this.recursiveSearch(id, this.Data.AtmGroupsGet.info.usersGroup[0].atmGroups));
            }
        }

        private GroupsGet.AtmGroup recursiveSearch(string id, List<GroupsGet.AtmGroup> info)
        {
            var result = new GroupsGet.AtmGroup();
            result.name = string.Empty;
            foreach (var group in info)
            {
                if (group.id.ToString() == id)
                    return group;
                if (!group.atmGroups.IsNullOrEmpty())
                {
                    result = recursiveSearch(id, group.atmGroups);
                    if (result.name != string.Empty)
                        return result;

                }
            }
            return result;
        }

        internal override void MakeAnExcel()
        {
            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\OUS_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

            this.CreateFailuresDict();

            this.SearchGroupNames();

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

                worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());


                Sheet sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    Name = ReportsSource.MonitoringService,
                    SheetId = (uint)1
                };
                sheets.Append(sheet);

                var gr = this.Data.AtmGroupsGet.info.usersGroup[0].atmGroups.Where(grr => this.Info.atmsGroupsId.Contains(grr.id.ToString())).ToList();
                var headGroups = groups.Where(gr1 => gr.Contains(gr1.Value)).ToList();

                CreateHeaderRow(worksheetPart);
                CreateFooterRow(worksheetPart);
                CreateTotalRows(worksheetPart, headGroups);
                CreateFooterRow(worksheetPart);

                this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportMonitoringColumn.xml");

                foreach (var group in headGroups)
                {                    
                    CreateNADRows(worksheetPart, group);
                    CreateNAARows(worksheetPart, group);
                    CreateCommRows(worksheetPart, group);
                    CreateEncashRows(worksheetPart, group);
                    CreateFooterRow(worksheetPart);
                }


                M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 1, 12);
                M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 2, 42);
                M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 3, 34);
                M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 4, 34);
                M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 5, 34);
                M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 6, 34);
                M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 7, 34);
                M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 8, 34);
                workbookpart.Workbook.Save();
            }
        }

        private bool SearchAtmId(string Id, GroupsGet.AtmGroup group)
        {
            var result = false;
            if (group.atmIds.Contains(Id))
                result = true;
            else
            {
                foreach (var gr in group.atmGroups)
                    if (SearchAtmId(Id, gr))
                        return true;               
            }                
            return result;
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = "Отчет о простаивающих банкоматах c " + this.Info.from + " по " + this.Info.to;
            for (int i = 1; i <= 7; i++)
            {
                if (i > 1)
                    title = "";
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A1", "H1");
        }

        private void CreateCommRows(WorksheetPart worksheetPart, KeyValuePair<string, GroupsGet.AtmGroup> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #region title
            string title = "Отсутствие связи";
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #endregion


            var incidents = this.Data.Incidents.Where(incident => (incident.deviceTypeId == failuresDict["Communications"] || incident.deviceTypeId == failuresDict["AgentComm"]) && SearchAtmId(incident.atmId, group.Value));

            #region head
            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {
                title = "Нет";
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                return;
            }

            #endregion

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.Where(atm => atm.Id == incident.atmId).First();

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, !String.IsNullOrEmpty(incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data)) 
                    ? incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data) 
                    : this.Data.DictionariesInfo.devicesTypes.data.First(device => device.id == Convert.ToInt32(incident.deviceTypeId)).description, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, date.ToString("yyyy-MM-dd hh:mm:ss"), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 7, row.RowIndex, incident.comments, CellValues.String, 5U);

            }

        }

        private void CreateNADRows(WorksheetPart worksheetPart, KeyValuePair<string, GroupsGet.AtmGroup> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #region title
            string title = group.Value.name;
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            title = "Ремонт.Недоступные на выдачу";
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            #endregion

            var incidents = this.Data.Incidents.Where(
                incident => (incident.deviceTypeId == failuresDict["CardReader"] ||
                             (incident.deviceTypeId == failuresDict["JournalPrinter"] && !incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data).ToUpper().Contains("БУМАГ"))) &&
                              SearchAtmId(incident.atmId, group.Value));

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {
                title = "Нет";
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.First(atm => atm.Id == incident.atmId);

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, !String.IsNullOrEmpty(incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data)) 
                    ? incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data) 
                    : this.Data.DictionariesInfo.devicesTypes.data.First(device => device.id == Convert.ToInt32(incident.deviceTypeId)).description, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                M3Utils.ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

        }

        private void CreateNAARows(WorksheetPart worksheetPart, KeyValuePair<string, GroupsGet.AtmGroup> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #region title
            var title = "Ремонт.Не доступныe на прием";
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            #endregion

            var incidents = this.Data.Incidents.Where(
                incident => (incident.deviceTypeId == failuresDict["BNA"] ||
                                                      (incident.deviceTypeId == failuresDict["ReceiptPrinter"] && !(incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data).ToUpper().Contains("БУМАГ")))) &&
                                                      SearchAtmId(incident.atmId, group.Value));

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {
                title = "Нет";
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.Where(atm => atm.Id == incident.atmId).First();

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, !String.IsNullOrEmpty(incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data))
                    ? incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data) 
                    : this.Data.DictionariesInfo.devicesTypes.data.Where(device => device.id == Convert.ToInt32(incident.deviceTypeId)).First().description, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                M3Utils.ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

        }

        private void CreateEncashRows(WorksheetPart worksheetPart, KeyValuePair<string, GroupsGet.AtmGroup> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #region title
            var title = "Инкассация";
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            #endregion
            var statuses = this.Data.DictionariesInfo.statuses.data.Where(stat => stat.text == "Инкассация").Select(stat => stat.id);

            var incidents = this.Data.Incidents.Where(incident => (statuses.Contains(Convert.ToInt32(incident.statusId))) && SearchAtmId(incident.atmId, group.Value));

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "G" + row.RowIndex);
                for (int i = 1; i <= this.reportColumns.Count; i++)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            else
            {
                title = "Нет";
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = this.Data.AtmInfo.Where(atm => atm.Id == incident.atmId).First();

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.DeviceNumber, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.Place, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, !String.IsNullOrEmpty(incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data))
                    ? incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data)
                    : this.Data.DictionariesInfo.devicesTypes.data.First(device => device.id == Convert.ToInt32(incident.deviceTypeId)).description, CellValues.String, 5U);

                M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.RecoveryTime, out hours))
                    date.AddHours(hours);

                M3Utils.ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

        }

        private void CreateTotalRows(WorksheetPart worksheetPart, List<KeyValuePair<string, GroupsGet.AtmGroup>> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            row = (Row)sheetData.LastChild;
            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;


            M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, "Сервис (Не работают на выдачу /работают на выдачу, но не работают на прием)", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, "Связь", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, "Инкассация", CellValues.String, 4U);


            foreach (var grp in group)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                var incAcc = this.Data.Incidents.Where(
               incident => (incident.deviceTypeId == failuresDict["BNA"] ||
                                                     (incident.deviceTypeId == failuresDict["ReceiptPrinter"] && !(incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data).ToUpper().Contains("БУМАГ")))) &&
                                                     SearchAtmId(incident.atmId, grp.Value)
                                                     ).ToList();
                var incDisp = this.Data.Incidents.Where(
                incident => (incident.deviceTypeId == failuresDict["CardReader"] ||
                                                      (incident.deviceTypeId == failuresDict["JournalPrinter"] && !(incident.GetSubject(this.Data.DictionariesInfo.incidentsRules.data).ToUpper().Contains("БУМАГ")))) &&
                                                       SearchAtmId(incident.atmId, grp.Value)
                                                      ).ToList();

                var comms = this.Data.Incidents.Where(incident => (incident.deviceTypeId == failuresDict["Communications"] || incident.deviceTypeId == failuresDict["AgentComm"]) && SearchAtmId(incident.atmId, grp.Value)).ToList();

                var statuses = this.Data.DictionariesInfo.statuses.data.Where(stat => stat.text == "Инкассация").Select(stat => stat.id);
                var encash = this.Data.Incidents.Where(incident => statuses.Contains(Convert.ToInt32(incident.statusId)) && SearchAtmId(incident.atmId, grp.Value)).ToList();


                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, grp.Value.name, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, incDisp.Count.ToString() + "/" + incAcc.Count.ToString(), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, comms.Count.ToString(), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, encash.Count.ToString(), CellValues.String, 5U);
            }
        }

        private void CreateFooterRow(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            row = (Row)sheetData.LastChild;
            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= 8; i++)
            {
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }
        }
    }
}