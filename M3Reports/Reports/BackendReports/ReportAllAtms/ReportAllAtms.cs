namespace M3Reports
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using M3Incidents;

    public class ReportAllAtms : ReportBuilder
    {
        private List<ReportColumns> reportColumns;

        internal override void MakeAnExcel()
        {
            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\ALL_ATMS_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

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
                    Name = ReportsSource.RegisteredAtms,
                    SheetId = (uint)1
                };
                sheets.Append(sheet);

                this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportAllAtmsColumn.xml");

                this.CreateHeaderRow(worksheetPart);
                this.CreateDataRows(worksheetPart);

                for (int i = 1; i <= this.reportColumns.Count(); i++)
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportColumns[i - 1].width);
            }
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart)
        {
            try
            {
                Row row;
                SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

                sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                string title = String.Join(" ", new [] { ReportsSource.ReportRegisteredAtmsOf, this.Info.to});

                for (int i = 1; i <= this.reportColumns.Count; i++)
                {
                    if (i > 1) title = "";

                    M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }

                M3Utils.ExcelHelper.MergeCellsInRange(
                    worksheetPart.Worksheet,
                    this.reportColumns[0].localtion + row.RowIndex,
                    this.reportColumns[this.reportColumns.Count - 1].localtion + row.RowIndex);

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });

                row = (Row)sheetData.LastChild;

                for (int i = 1; i <= this.reportColumns.Count; i++)
                    M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            catch (Exception exp)
            {
                M3Utils.Log.Instance.Info(
                    this + ".CreateHeaderRow(...) exception:",
                    exp.Message,
                    exp.Source,
                    exp.StackTrace);
            }
        }

        private void CreateDataRows(WorksheetPart worksheetPart)
        {
            try
            {
                Row row;
                SheetData sheetData;

                sheetData = (SheetData)worksheetPart.Worksheet.First();
                row = (Row)sheetData.LastChild;

                for (int i = 0; i < this.Data.AtmInfo.Count; i++)
                {
                    sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                    row = (Row)sheetData.LastChild;

                    List<Incident> actualIncidents = this.Data.Incidents.Where(inc => inc.atmId == this.Data.AtmInfo[i].Id).OrderBy(inc => DateTime.Parse(inc.timeCreated)).ToList();

                    M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, this.Data.AtmInfo[i].DeviceNumber, CellValues.String, 2U);
                    M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, this.Data.AtmInfo[i].GeoAddress, CellValues.String, 2U);
                    M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, this.Data.AtmInfo[i].Sn, CellValues.String, 2U);
                    M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, String.Join(", ", (from item in this.Data.AtmGroups where item.atmIds.Contains(this.Data.AtmInfo[i].Id) select item.name).ToArray()), CellValues.String, 2U);
                    M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, (actualIncidents.Count == 0) ? ReportsSource.InService : ReportsSource.OutOfService, CellValues.String, 2U);
                    M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, (actualIncidents.Count != 0) ? this.Data.DictionariesGet.UserRoles.Where(role => role.id == actualIncidents.Last().userRoleId).ToArray()[0].description : string.Empty, CellValues.String, 2U);
                }
            }
            catch (Exception exp)
            {
                M3Utils.Log.Instance.Info(
                    this + ".CreateDataRows(...) exception:",
                    exp.Message,
                    exp.Source,
                    exp.StackTrace);
            }
        }
    }
}