using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using M3Atms;
using M3Incidents;
using M3Dictionaries;

namespace M3Reports
{
    public class ReportIncidentsPivotTable : ReportBuilder
    {
        public List<ReportColumns> reportColumns = new List<ReportColumns>();

        internal override void MakeAnExcel()
        {
            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\IPT_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

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
                    Name = "Сводная таблица инцидентов",
                    SheetId = 1
                };

                sheets.Append(sheet);

                this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportIncidentsPivotTableColumn.xml");

                this.CreateHeaderRow(worksheetPart);
                this.CreateDataRows(worksheetPart);

                for (int i = 1; i <= this.reportColumns.Count(); i++)
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportColumns[i - 1].width);

                workbookpart.Workbook.Save();
            }
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = ReportsSource.ReportIncidentTable;

            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                if (i > 1)
                    title = "";

                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportColumns[0].localtion + row.RowIndex, this.reportColumns[this.reportColumns.Count - 1].localtion + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });

            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportColumns.Count; i++)
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
        }

        private void CreateDataRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData;

            sheetData = (SheetData)worksheetPart.Worksheet.First();
            row = (Row)sheetData.LastChild;

            for (int i = 0; i < this.Info.incidents.Count; i++)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                Info atmInfo = this.GetAtmInfoByAtmId(this.Info.incidents[i].atmId);

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, this.Info.incidents[i].timeCreated.Replace("-", "").Substring(2, 6) + this.Info.incidents[i].id, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, ((atmInfo != null) ? atmInfo.DeviceNumber : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, ((atmInfo != null) ? atmInfo.Vizname : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, ((atmInfo != null) ? atmInfo.Model : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, ((atmInfo != null) ? atmInfo.Sn : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, ((atmInfo != null) ? atmInfo.City : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 7, row.RowIndex, ((atmInfo != null) ? atmInfo.GeoAddress : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 8, row.RowIndex, ((atmInfo != null) ? atmInfo.Place : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 9, row.RowIndex, this.GetStatusById(Convert.ToInt32(this.Info.incidents[i].statusId)), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 10, row.RowIndex, this.Info.incidents[i].comments, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 11, row.RowIndex, this.GetBankDivisionById(this.Info.incidents[i].userRoleId), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 12, row.RowIndex, this.GetUserById(Convert.ToInt32(this.Info.incidents[i].assignedToId)), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 13, row.RowIndex, this.GetResponsibleForId(this.Info.incidents[i].responsibleForId), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 14, row.RowIndex, this.Info.incidents[i].timeCreated, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 15, row.RowIndex, this.Info.incidents[i].timeProceedToService, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 16, row.RowIndex, ((this.Info.incidents[i].isCritical == 1) ? ReportsSource.Yes : ReportsSource.No), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 17, row.RowIndex, ((atmInfo != null) ? atmInfo.WorkHours : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 18, row.RowIndex, ((atmInfo != null) ? atmInfo.EncashHours : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 19, row.RowIndex, ((atmInfo != null) ? atmInfo.AdviceSum : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 20, row.RowIndex, string.Empty, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 21, row.RowIndex, string.Empty, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 22, row.RowIndex, ((this.Info.incidents[i].authorId == "0") ? "M3 Incident Manager" : this.GetUserById(Convert.ToInt32(this.Info.incidents[i].authorId))), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 23, row.RowIndex, ((this.Info.incidents[i].timeChanged.Any()) ? this.Info.incidents[i].timeChanged : this.Info.incidents[i].timeCreated), CellValues.String, 5U); 
            }
        }

        private Info GetAtmInfoByAtmId(string id)
        {
            List<Info> atmInfoList = (from item in this.Data.AtmInfo
                                         where item.Id == id
                                         select item).ToList();

            return (atmInfoList.Count > 0) ? atmInfoList.First() : null;
        }

        private string GetStatusById(int id)
        {
            List<string> statusList = (from item in this.Data.DictionariesGet.Statuses
                                       where item.id == id
                                       select item.text).ToList();

            return (statusList.Count > 0) ? statusList.First() : "";
        }

        private string GetBankDivisionById(int id)
        {
            List<string> bankDivisionList = (from item in this.Data.DictionariesGet.UserRoles
                                             where item.id == id
                                             select item.description).ToList();

            return (bankDivisionList.Count > 0) ? bankDivisionList.First() : "";
        }

        private string GetUserById(int id)
        {
            List<string> userList = (from item in this.Data.DictionariesGet.Users
                                     where item.id == id
                                     select item.lName + " " + item.lName).ToList();

            return (userList.Count > 0) ? userList.First() : "";
        }

        private string GetResponsibleForId(string id)
        {
            List<string> responsibleForList = (from item in this.Data.DictionariesGet.ResponsibleFor
                                               where item.id == id
                                               select item.text).ToList();

            return (responsibleForList.Count > 0) ? responsibleForList.First() : "";
        }
    }
}