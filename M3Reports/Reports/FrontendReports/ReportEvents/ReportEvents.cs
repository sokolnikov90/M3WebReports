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
    public class ReportEvents : ReportBuilder
    {
        public List<EventItem> eventItems;
        public List<EventDescriptionItem> eventDescriptionItems;

        public List<ReportColumns> reportColumns { get; set; }

        public void ParseEvents(XmlNode messageNode)
        {
            XmlNodeList eventXmlList = messageNode.SelectNodes("Request/Events/Event");

            EventItem eventItem;

            if (eventXmlList != null)
            {
                this.eventItems = new List<EventItem>();

                for (int i = 0; i < eventXmlList.Count; i++)
                {
                    eventItem = new EventItem()
                    {
                        Id = eventXmlList[i].SelectSingleNode("Id").InnerText,
                        DTime = eventXmlList[i].SelectSingleNode("DateTime").InnerText,
                        AtmId = eventXmlList[i].SelectSingleNode("AtmId").InnerText
                    };

                    this.eventItems.Add(eventItem);
                }
            }
        }

        public void ParseEventsDescriptions(XmlNode messageNode)
        {
            XmlNodeList eventDescriptionXmlList = messageNode.SelectNodes("Request/Events/Event");

            EventDescriptionItem eventDescriptionItem;

            if (eventDescriptionXmlList != null)
            {
                this.eventDescriptionItems = new List<EventDescriptionItem>();

                for (int i = 0; i < eventDescriptionXmlList.Count; i++)
                {
                    eventDescriptionItem = new EventDescriptionItem()
                    {
                        Id = eventDescriptionXmlList[i].SelectSingleNode("Id").InnerText,
                        Name = eventDescriptionXmlList[i].SelectSingleNode("Name").InnerText
                    };

                    this.eventDescriptionItems.Add(eventDescriptionItem);
                }
            }
        }

        internal override void MakeAnExcel()
        {
            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\EVT_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

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
                    Name = ReportsSource.Events,
                    SheetId = 1
                };

                sheets.Append(sheet);

                this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportEventsColumn.xml");

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

            string title = String.Join(" ", new[]
                                    {
                                        ReportsSource.ReportOnTheEventsAtTheAtm,
                                        this.Data.AtmInfo[0].Vizname,
                                        this.Data.AtmInfo[0].GeoAddress,
                                        ReportsSource.From,
                                        this.Info.from,
                                        ReportsSource.To,
                                        this.Info.to
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
            {
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
        }

        private void CreateDataRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            string description;
            row = (Row)sheetData.LastChild;

            foreach (EventItem eventItem in this.eventItems)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                description = this.eventDescriptionItems.First(descr => descr.Id == eventItem.Id).Name;
                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, description, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, eventItem.DTime, CellValues.String, 5U);
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                M3Utils.ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }
        }
    }
}