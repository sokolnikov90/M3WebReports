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
    public class ReportIncidentsOpen : ReportBuilder
    {
        private List<ReportColumns> reportColumns = new List<ReportColumns>();

        internal override void MakeAnExcel()
        {
            try
            {
                string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

                this.Info.path = this.Info.path.Replace("/", "\\") + "\\RIO_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

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

                    this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportIncidentsOpenColumn.xml");

                    for (int i = 0; i < this.Data.DictionariesGet.UserRoles.Count; i++)
                    {
                        if (this.Data.DictionariesGet.UserRoles[i].description == ReportsSource.UserOfM3WEB)
                            continue;

                        worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());

                        Sheet sheet = new Sheet()
                        {
                            Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                            Name = this.Data.DictionariesGet.UserRoles[i].description.Length <= 22 ?
                                this.Data.DictionariesGet.UserRoles[i].description :
                                this.Data.DictionariesGet.UserRoles[i].description.Substring(0,22) + "...",
                            SheetId = Convert.ToUInt32(i + 1)
                        };

                        sheets.Append(sheet);

                        this.CreateHeaderRow(worksheetPart, this.Info.from, this.Info.to);
                        this.CreateDataRows(worksheetPart, this.Data.DictionariesGet.UserRoles[i].id);

                        for (int j = 1; j <= this.reportColumns.Count(); j++)
                            M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, j, this.reportColumns[j - 1].width);
                    }

                    workbookpart.Workbook.Save();
                }
            }
            catch (Exception exp)
            {
                M3Utils.Log.Instance.Info(this + ".MakeAnExcel() exeption: " + exp.Message);
            }
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart, string from, string to)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = string.Join(" ", new[]
                        {
                            ReportsSource.IncidentReport,
                            ReportsSource.From, this.Info.from,
                            ReportsSource.To, this.Info.to
                        }); 

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

        private void CreateDataRows(WorksheetPart worksheetPart, int userRoleId)
        {
            Row row;
            SheetData sheetData;

            sheetData = (SheetData)worksheetPart.Worksheet.First();
            row = (Row)sheetData.LastChild;

            List<Incident> selectedIncidents = (from item in this.Data.Incidents
                                                where item.userRoleId == userRoleId
                                                select item).ToList();

            for (int i = 0; i < selectedIncidents.Count; i++)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                Info atmInfo = this.GetAtmInfoByAtmId(selectedIncidents[i].atmId);

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, selectedIncidents[i].timeCreated.Replace("-", "").Substring(2, 6) + selectedIncidents[i].id, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, ((atmInfo != null) ? atmInfo.GeoAddress : ""), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, this.GetTypeById(Convert.ToInt32(selectedIncidents[i].typeId)), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, selectedIncidents[i].comments, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, ((int)(DateTime.Now - DateTime.Parse(selectedIncidents[i].timeCreated)).TotalHours).ToString(), CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, ((selectedIncidents[i].isCritical == 1) ? ReportsSource.Yes : ReportsSource.No), CellValues.String, 5U);
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

        private string GetResponsibleForId(int id)
        {
            List<string> responsibleForList = (from item in this.Data.DictionariesGet.ResponsibleFor
                                               where Convert.ToInt32(item.id) == id
                                               select item.text).ToList();

            return (responsibleForList.Count > 0) ? responsibleForList.First() : "";
        }

        private string GetTypeById(int id)
        {
            List<string> typeList = (from item in this.Data.DictionariesGet.Types
                                               where item.id == id
                                               select item.text).ToList();

            return (typeList.Count > 0) ? typeList.First() : "";
        }
    }
}