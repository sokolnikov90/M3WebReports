using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

using M3Atms;
using M3Incidents;
using M3Dictionaries;

namespace M3Reports
{
    public class ReportSummaryIncidentsOpen : ReportBuilder
    {
        public static readonly Dictionary<string, List<int>> PROBLEMS_DICTIONARY = new Dictionary<string, List<int>>
        {
            {"Не обработан", new List<int>(){ 1 }},                                         // Открыт
            {"SLM", new List<int>(){ 4 }},                                     // Ремонт
            {"Не работает. Канал связи", new List<int>(){ 5, 6 }},                          // Проблемная связь, отсутсвие связи
            {"Не работает. Проблема с электропитанием", new List<int>(){ 7, 8 }},           // Проблемное электропитание, нет электропитания
            {"FLM", new List<int>(){ 9, 16, 20 }},                  // Технический выезд, замятие/замена бумаги в принтере, совместный выезд
            {"Не работает. Проблема с сигнализацией", new List<int>(){ 10 }},               // Ремонт/установка сигнализации
            {"Демонтаж/перемещение", new List<int>(){ 11 }},                                // Демонтаж/перемещение
            {"Изменение номинала", new List<int>(){ 12 }},                                  // Изменение номинала
            {"Не работает. Инкассация. Загрузка, выгрузка", new List<int>(){ 13, 14, 15 }}, // Переинкассация, загрузка, разгрузка
            {"Отсутсвие транзакций", new List<int>(){ 17 }},                                // Отсутсвие транзакций
            {"Проблемы с доступом к УС", new List<int>(){ 18 }},                            // Доступ к УС
            {"Закончились деньги", new List<int>(){ 19 }}                                   // Деньги заканчиваются
        };

        public int rootGroupsCount;

        public List<ReportColumns> reportPage1Columns = new List<ReportColumns>();
        public List<ReportColumns> reportPageNColumns = new List<ReportColumns>();
        public List<ReportColumns> reportPageAtmsColumns = new List<ReportColumns>();

        public M3Dictionaries.GetAll.Info dictionariesInfo;

        internal override void MakeAnExcel()
        {
            try
            {
                string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

                StringBuilder path = new StringBuilder();

                switch (this.Info.type)
                {
                    case "SummaryIncidentsOpen":
                        path.AppendFormat("{0}\\RSIO_{1}{2}_{3}{4}.xlsx", this.Info.path.Replace("/", "\\"), fromArray[0].Replace("-", "").Substring(2), fromArray[1].Replace(":", ""), toArray[0].Replace("-", "").Substring(2), toArray[1].Replace(":", ""));
                        break;
                    case "SummaryIncidentsOpenWorking":
                        path.AppendFormat("{0}\\RSIO_W_{1}{2}_{3}{4}.xlsx", this.Info.path.Replace("/", "\\"), fromArray[0].Replace("-", "").Substring(2), fromArray[1].Replace(":", ""), toArray[0].Replace("-", "").Substring(2), toArray[1].Replace(":", ""));
                        break;
                }

                this.Info.path = path.ToString();

                this.reportPage1Columns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportSummaryIncidentsOpenPage1Column.xml");
                this.reportPageAtmsColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportAllAtmsColumn.xml");
                this.reportPageNColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportSummaryIncidentsOpenPageNColumn.xml");

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
                    
                    #region Sheet#1

                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    Sheet sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                        Name = "Итого",
                        SheetId = Convert.ToUInt32(1)
                    };

                    sheets.Append(sheet);

                    this.CreatePage1HeaderRow(worksheetPart);
                    this.CreatePage1DataRows(worksheetPart);

                    for (int i = 1; i <= this.reportPage1Columns.Count(); i++)
                        M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportPage1Columns[i - 1].width);

                    for (int i = 1; i <= this.rootGroupsCount; i++)
                        M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i + this.reportPage1Columns.Count(), 20.0f);

                    #endregion

                    #region Sheet#2

                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                        Name = "Работает",
                        SheetId = (uint)2
                    };

                    sheets.Append(sheet);

                    this.CreateHeaderRowAtms(worksheetPart);
                    this.CreateDataRowsAtms(worksheetPart, true);

                    for (int i = 1; i <= this.reportPageAtmsColumns.Count(); i++)
                        M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportPageAtmsColumns[i - 1].width);
                    
                    #endregion

                    #region Sheet#3

                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                        Name = "Зарегистрировано",
                        SheetId = (uint)3
                    };

                    sheets.Append(sheet);

                    this.CreateHeaderRowAtms(worksheetPart);
                    this.CreateDataRowsAtms(worksheetPart, false);

                    for (int i = 1; i <= this.reportPageAtmsColumns.Count(); i++)
                        M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportPageAtmsColumns[i - 1].width);

                    #endregion

                    #region Sheet#N

                    for (int i = 0; i < this.dictionariesInfo.userRoles.data.Count; i++)
                    {
                        if (this.dictionariesInfo.userRoles.data[i].description == "Пользователь M3 WEB")
                            continue;

                        worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());

                        Sheet sheetN = new Sheet()
                        {
                            Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                            Name = this.dictionariesInfo.userRoles.data[i].description,
                            SheetId = Convert.ToUInt32(i + 4)
                        };

                        sheets.Append(sheetN);

                        this.CreatePageNHeaderRow(worksheetPart, this.dictionariesInfo.userRoles.data[i]);
                        this.CreatePageNDataRows(worksheetPart, this.dictionariesInfo.userRoles.data[i]);

                        for (int j = 1; j <= this.reportPageNColumns.Count(); j++)
                            M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, j, this.reportPageNColumns[j - 1].width);
                    }

                    #endregion

                    workbookpart.Workbook.Save();
                }
            }
            catch (Exception exp)
            {
                M3Utils.Log.Instance.Info(this + ".MakeAnExcel() exeption: " + exp.Message);
            }
        }

        private void CreatePage1HeaderRow(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 60D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            StringBuilder title = new StringBuilder();
            switch (this.Info.type)
            {
                case "SummaryIncidentsOpen":
                    title.AppendFormat("Сводный отчет по инцидентам\n Данные на: {0}\n Отчет создан: {1}", this.Info.to, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    break;
                case "SummaryIncidentsOpenWorking":
                    title.AppendFormat("Сводный отчет по инцидентам. Рабочие инциденты\n Данные на: {0}\n Отчет создан: {1}", this.Info.to, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    break;
            }

            for (int i = 0; i < this.Data.AtmGroups.Count; i++)
            {
                if (this.Info.atmsGroupsId.Contains(this.Data.AtmGroups[i].id.ToString())) this.rootGroupsCount++;
            }

            for (int i = 1; i <= (this.reportPage1Columns.Count + this.rootGroupsCount); i++)
            {
                if (i > 1)
                {
                    title.Length = 0;
                    title.Capacity = 0;
                }

                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title.ToString(), CellValues.String, 4U);
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportPage1Columns[0].localtion + row.RowIndex, M3Utils.ExcelHelper.ColumnNameByIndex(this.reportPage1Columns.Count + this.rootGroupsCount) + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 2), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportPage1Columns.Count; i++)
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportPage1Columns[i - 1].title, CellValues.String, 4U);

            for (int i = 0, j = 0; i < this.Data.AtmGroups.Count; i++)
            {
                if (this.Info.atmsGroupsId.Contains(this.Data.AtmGroups[i].id.ToString()))
                {
                    M3Utils.ExcelHelper.CreateCell(row, (this.reportPage1Columns.Count + 1) + j, row.RowIndex, this.Data.AtmGroups[i].name, CellValues.String, 4U);
                    j++;
                }
            }
        }

        public void CreatePage1DataRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData;

            sheetData = (SheetData)worksheetPart.Worksheet.First();
            row = (Row)sheetData.LastChild;

            for (int i = 0; i < 2; i++)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, ReportsSource.Total, CellValues.String, 2U);

                if (i == 0)
                {
                    M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, ReportsSource.Registered, CellValues.String, 2U);
                    M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, this.Info.atmsId.Count().ToString(), CellValues.String, 2U);

                    for (int k = 0, l = 0; k < this.Data.AtmGroups.Count; k++)
                    {
                        if (this.Info.atmsGroupsId.Contains(Convert.ToString(this.Data.AtmGroups[k].id)))
                        {
                            var atmsInGroupAndInReport = this.Data.AtmGroups[k].atmIds.Intersect(this.Info.atmsId);

                            M3Utils.ExcelHelper.CreateCell(row, (this.reportPage1Columns.Count + 1) + l, row.RowIndex, atmsInGroupAndInReport.Count().ToString(), CellValues.String, 2U);
                            l++;
                        }
                    }
                }

                if (i == 1)
                {
                    List<string> problemsDictionaryKeys;
                    List<Incident> selectedProblemIncidents;

                    problemsDictionaryKeys = PROBLEMS_DICTIONARY.Keys.ToList();

                    selectedProblemIncidents = (from item in this.Data.Incidents
                                                where PROBLEMS_DICTIONARY.Values.SelectMany(x => x).Contains(Convert.ToInt32(item.statusId))
                                                select item).ToList();

                    M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, ReportsSource.Working, CellValues.String, 2U);

                    for (int k = 0, l = 0; k < this.Data.AtmGroups.Count; k++)
                    {
                        if (this.Info.atmsGroupsId.Contains(Convert.ToString(this.Data.AtmGroups[k].id)))
                        {
                            var atmsInGroupAndInReport = this.Data.AtmGroups[k].atmIds.Intersect(this.Info.atmsId);

                            var selectedIncidentIdsInGroup = (from item in selectedProblemIncidents
                                                              where this.Data.AtmGroups[k].atmIds.Contains(item.atmId)
                                                              select item.atmId).Intersect(this.Info.atmsId);

                            if (selectedIncidentIdsInGroup.Count() == 0)
                                M3Utils.ExcelHelper.CreateCell(row, (this.reportPage1Columns.Count + 1) + l, row.RowIndex, atmsInGroupAndInReport.Count().ToString(), CellValues.String, 2U);
                            else
                                M3Utils.ExcelHelper.CreateCell(row, (this.reportPage1Columns.Count + 1) + l, row.RowIndex, (atmsInGroupAndInReport.Count() - selectedIncidentIdsInGroup.Count()).ToString(), CellValues.String, 2U);

                            l++;
                        }
                    }

                    if (selectedProblemIncidents.Count == 0)
                        M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, this.Info.atmsId.Count().ToString(), CellValues.String, 2U);
                    else
                        M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, (this.Info.atmsId.Count() - selectedProblemIncidents.GroupBy(p => p.atmId).Count()).ToString(), CellValues.String, 2U);
                }
            }

            for (int i = 0; i < this.dictionariesInfo.userRoles.data.Count; i++)
            {
                List<Incident> selectedIncidents = (from item in this.Data.Incidents
                                                    where ((item.userRoleId == this.dictionariesInfo.userRoles.data[i].id) && (this.GetStatusById(Convert.ToInt32(item.statusId)) != "Закрыт"))
                                                    select item).ToList();

                if (selectedIncidents.Count == 0)
                    continue;

                List<string> problemsDictionaryKeys;
                List<Incident> selectedProblemIncidents;

                problemsDictionaryKeys = PROBLEMS_DICTIONARY.Keys.ToList();

                for (int j = 0; j < problemsDictionaryKeys.Count; j++)
                {
                    selectedProblemIncidents = (from item in selectedIncidents
                                                where PROBLEMS_DICTIONARY[problemsDictionaryKeys[j]].Contains(Convert.ToInt32(item.statusId))
                                                select item).ToList();

                    if (selectedProblemIncidents.Count == 0)
                        continue;

                    sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                    row = (Row)sheetData.LastChild;

                    M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, this.dictionariesInfo.userRoles.data[i].description, CellValues.String, 2U);
                    M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, problemsDictionaryKeys[j], CellValues.String, 2U);
                    M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, selectedProblemIncidents.GroupBy(p => p.atmId).Select(g => g.First()).ToList().Count.ToString(), CellValues.String, 2U);

                    for (int k = 0, l = 0; k < this.Data.AtmGroups.Count; k++)
                    {
                        if (this.Info.atmsGroupsId.Contains(Convert.ToString(this.Data.AtmGroups[k].id)))
                        {
                            var atmsInGroupAndInReport = this.Data.AtmGroups[k].atmIds.Intersect(this.Info.atmsId);

                            List<Incident> selectedIncidentsInGroup = (from item in selectedProblemIncidents
                                                                       where this.Data.AtmGroups[k].atmIds.Contains(item.atmId)
                                                                       select item).ToList();

                            M3Utils.ExcelHelper.CreateCell(row, (this.reportPage1Columns.Count + 1) + l, row.RowIndex, selectedIncidentsInGroup.GroupBy(p => p.atmId).Select(g => g.First()).ToList().Count.ToString(), CellValues.String, 2U);
                            l++;
                        }
                    }
                }
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;

            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, ReportsSource.Availability, CellValues.String, 5U);

            List<string> brokenAtms = (from item in this.Data.Incidents
                                       where ((this.Info.atmsId.Contains(item.atmId)) && (item.isCritical == 1) && (this.GetStatusById(Convert.ToInt32(item.statusId)) != "Закрыт"))
                                       select item.atmId).Distinct().ToList();

            double amountAvailability = (this.Info.atmsId.Any()) ? (double)100 * (this.Info.atmsId.Count() - brokenAtms.Count()) / this.Info.atmsId.Count() : 100.0f;

            UInt32 styleIndex;

            if (amountAvailability >= 95.0f)
                styleIndex = 7U;
            else if (amountAvailability >= 80.0f && amountAvailability < 95.0f)
                styleIndex = 8U;
            else
                styleIndex = 9U;

            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, string.Format("{0:F2}%", amountAvailability), CellValues.String, styleIndex);

            for (int i = 0, j = 0; i < this.Data.AtmGroups.Count; i++)
            {
                if (this.Info.atmsGroupsId.Contains(this.Data.AtmGroups[i].id.ToString()))
                {
                    var atmsInGroupAndInReport = this.Data.AtmGroups[i].atmIds.Intersect(this.Info.atmsId);

                    var brokenAtmsInGroup = (from item in this.Data.Incidents
                                             where (this.Data.AtmGroups[i].atmIds.Contains(item.atmId) && (item.isCritical == 1) && (this.GetStatusById(Convert.ToInt32(item.statusId)) != "Закрыт"))
                                             select item.atmId).Distinct();

                    amountAvailability = (atmsInGroupAndInReport.Any()) ? (double)100 * (atmsInGroupAndInReport.Count() - brokenAtmsInGroup.Count()) / atmsInGroupAndInReport.Count() : 100.0f;

                    if (amountAvailability >= 95.0f)
                        styleIndex = 7U;
                    else if (amountAvailability >= 80.0f && amountAvailability < 95.0f)
                        styleIndex = 8U;
                    else
                        styleIndex = 9U;

                    M3Utils.ExcelHelper.CreateCell(row, (this.reportPage1Columns.Count + 1) + j, row.RowIndex, string.Format("{0:F2}%", amountAvailability), CellValues.String, styleIndex);
                    j++;
                }
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 2) });
            row = (Row)sheetData.LastChild;

            string title = ReportsSource.More + " 95%";

            for (int i = 1; i <= this.reportPage1Columns.Count + this.rootGroupsCount; i++)
            {
                if (i > 1)
                    title = "";

                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 7U);
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportPage1Columns[0].localtion + row.RowIndex, M3Utils.ExcelHelper.ColumnNameByIndex(this.reportPage1Columns.Count + this.rootGroupsCount) + row.RowIndex);
            //M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, "Более 95%", CellValues.String, 7U);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;

            title = string.Join(" ", new []
                                         {
                                             ReportsSource.From,
                                             "80%",
                                             ReportsSource.To,
                                             "95%"
                                         });

            for (int i = 1; i <= this.reportPage1Columns.Count + this.rootGroupsCount; i++)
            {
                if (i > 1)
                    title = "";

                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 8U);
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportPage1Columns[0].localtion + row.RowIndex, M3Utils.ExcelHelper.ColumnNameByIndex(this.reportPage1Columns.Count + this.rootGroupsCount) + row.RowIndex);
            //M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, "От 80% до 95%", CellValues.String, 8U);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;

            title = ReportsSource.Less + " 80%";

            for (int i = 1; i <= this.reportPage1Columns.Count + this.rootGroupsCount; i++)
            {
                if (i > 1)
                    title = "";

                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 9U);
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportPage1Columns[0].localtion + row.RowIndex, M3Utils.ExcelHelper.ColumnNameByIndex(this.reportPage1Columns.Count + this.rootGroupsCount) + row.RowIndex);
            //M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, "Менее 80%", CellValues.String, 9U);
        }

        private void CreateHeaderRowAtms(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 60D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = string.Empty;

            switch (this.Info.type)
            {
                case "SummaryIncidentsOpen":
                    title = String.Join(" ", new[]
                                        {
                                            ReportsSource.ReportSummaryIncidents,
                                            ReportsSource.of,
                                            this.Info.to
                                        });
                       break;
                case "SummaryIncidentsOpenWorking":
                       title = String.Join(" ", new[]
                                        {
                                            ReportsSource.ReportSummaryIncidentsWorkingIncidents,
                                            ReportsSource.of,
                                            this.Info.to
                                        });
                        break;
            }

            for (int i = 1; i <= this.reportPageAtmsColumns.Count; i++)
            {
                if (i > 1)
                {
                    title = string.Empty;
                }

                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportPageAtmsColumns[0].localtion + row.RowIndex, this.reportPageAtmsColumns[this.reportPageAtmsColumns.Count - 1].localtion + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 2), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportPageAtmsColumns.Count; i++)
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportPageAtmsColumns[i - 1].title, CellValues.String, 4U);
        }

        public void CreateDataRowsAtms(WorksheetPart worksheetPart, Boolean isWorking)
        {
            Row row;
            SheetData sheetData;

            sheetData = (SheetData)worksheetPart.Worksheet.First();
            row = (Row)sheetData.LastChild;

            for (int i = 0; i < this.Data.AtmInfo.Count; i++)
            {
                List<Incident> actualIncidents = this.Data.Incidents.Where(inc => (inc.atmId == this.Data.AtmInfo[i].Id)).OrderBy(inc => DateTime.Parse(inc.timeCreated)).ToList();

                //Если необхомо вывести работающие банкоматы, а количество инцидентов на текущем банкомате != 0
                if (isWorking && (actualIncidents.Count != 0))
                    continue;

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, this.Data.AtmInfo[i].DeviceNumber, CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, this.Data.AtmInfo[i].GeoAddress, CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, this.Data.AtmInfo[i].Sn, CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, this.GetGroupForAtm(this.Data.AtmInfo[i].Id), CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, (actualIncidents.Count == 0) ? "Обслуживает" : "Не обслуживает", CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, (actualIncidents.Count != 0) ? this.dictionariesInfo.userRoles.data.Where(role => role.id == actualIncidents.Last().userRoleId).ToArray()[0].description : "-", CellValues.String, 2U);
            }
        }

        private void CreatePageNHeaderRow(WorksheetPart worksheetPart, GetUserRoles.Data userRole)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 60D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = "";

            switch (this.Info.type)
            {
                case "SummaryIncidentsOpen":
                    title = String.Join(" ", new[]
                                        {
                                            ReportsSource.ReportSummaryIncidents,
                                            ReportsSource.of,
                                            this.Info.to
                                        });
                    break;
                case "SummaryIncidentsOpenWorking":
                    title = String.Join(" ", new[]
                                        {
                                            ReportsSource.ReportSummaryIncidentsWorkingIncidents,
                                            ReportsSource.of,
                                            this.Info.to
                                        });
                    break;
            }

            for (int i = 1; i <= this.reportPageNColumns.Count; i++)
            {
                if (i > 1)
                    title = "";

                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportPageNColumns[0].localtion + row.RowIndex, this.reportPageNColumns[this.reportPageNColumns.Count - 1].localtion + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 2), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportPageNColumns.Count; i++)
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportPageNColumns[i - 1].title, CellValues.String, 4U);
        }

        public void CreatePageNDataRows(WorksheetPart worksheetPart, GetUserRoles.Data userRole)
        {
            Row row;
            SheetData sheetData;

            sheetData = (SheetData)worksheetPart.Worksheet.First();
            row = (Row)sheetData.LastChild;

            List<Incident> selectedIncidents = new List<Incident>();

            //В incidents хранятся только рабочие и закрытые инциденты.
            switch (this.Info.type)
            {
                case "SummaryIncidentsOpen":
                    selectedIncidents = (from item in this.Data.Incidents
                                         where item.userRoleId == userRole.id
                                         select item).ToList();
                    break;
                case "SummaryIncidentsOpenWorking":
                    selectedIncidents = (from item in this.Data.Incidents
                                         where item.userRoleId == userRole.id
                                         select item).ToList();
                    break;
            }

            for (int i = 0; i < selectedIncidents.Count; i++)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                Info atmInfo = this.GetAtmInfoByAtmId(Convert.ToInt32(selectedIncidents[i].atmId));

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, selectedIncidents[i].timeCreated.Replace("-", "").Substring(2, 6) + selectedIncidents[i].id, CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, this.GetStatusById(Convert.ToInt32(selectedIncidents[i].statusId)), CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, ((atmInfo != null) ? atmInfo.DeviceNumber : ""), CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, ((atmInfo != null) ? atmInfo.Vizname : ""), CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, selectedIncidents[i].timeCreated, CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, selectedIncidents[i].timeProceedToService, CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 7, row.RowIndex, selectedIncidents[i].comments, CellValues.String, 2U);
                M3Utils.ExcelHelper.CreateCell(row, 8, row.RowIndex, ((selectedIncidents[i].isCritical == 1) ? ReportsSource.Yes : ReportsSource.No), CellValues.String, 2U);
            }
        }

        private GroupsGet.AtmGroup FindGroupById(string id, List<GroupsGet.AtmGroup> atmGroups)
        {
            GroupsGet.AtmGroup atmGroup = new GroupsGet.AtmGroup() { id = -1 };

            for (int i = 0; i < atmGroups.Count; i++)
            {
                if (atmGroups[i].id == Convert.ToInt32(id))
                {
                    atmGroup = atmGroups[i];
                    break;
                }

                if (atmGroups[i].atmGroups.Count > 0)
                    atmGroup = this.FindGroupById(id, atmGroups[i].atmGroups);

                if (atmGroup.id >= 0)
                    break;
            }

            return atmGroup;
        }

        private Info GetAtmInfoByAtmId(int id)
        {
            List<Info> atmInfoList = (from item in this.Data.AtmInfo
                                      where Convert.ToInt32(item.Id) == id
                                      select item).ToList();

            return (atmInfoList.Count > 0) ? atmInfoList.First() : null;
        }

        private string GetStatusById(int id)
        {
            List<string> statusList = (from item in this.dictionariesInfo.statuses.data
                                       where item.id == Convert.ToInt32(id)
                                       select item.text).ToList();

            return (statusList.Count > 0) ? statusList.First() : "";
        }

        private string GetBankDivisionById(int id)
        {
            List<string> bankDivisionList = (from item in this.dictionariesInfo.userRoles.data
                                             where item.id == id
                                             select item.description).ToList();

            return (bankDivisionList.Count > 0) ? bankDivisionList.First() : "";
        }

        private string GetGroupForAtm(string sAtmId)
        {
            try
            {
                for (int k = 0; k < this.Data.AtmGroups.Count; k++)
                {
                    if (this.Data.AtmGroups[k].atmIds.Contains(sAtmId))
                    {
                        return this.Data.AtmGroups[k].name;
                    }
                }
            }
            catch (Exception e)
            {
                M3Utils.Log.Instance.Info("GetBusinessGroupForAtm ReportAvailabilities(...) exception:");
                M3Utils.Log.Instance.Info(e.Message);
                M3Utils.Log.Instance.Info(e.Source);
                M3Utils.Log.Instance.Info(e.StackTrace);
            }
            return ReportsSource.Unknown;
        }
    }
}