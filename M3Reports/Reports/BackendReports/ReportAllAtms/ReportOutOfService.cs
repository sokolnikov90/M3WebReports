using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace M3WebService
{
    public class ReportOutOfService
    {
        public object Info
        {
            get;
            set;
        }

        public struct ReportOutOfServiceData
        {
            public List<AtmInfo> AtmInfoLst;
            public List<Incident> ActualIncidents;
        }

        public struct ReportOutOfServiceInfo
        {
            public int isError;
            public string from;
            public string to;
            public string path;
            public ReportOutOfServiceData data;
        }

        public object IsError { get; set; }
        public object From { get; set; }
        public object To { get; set; }
        public object Path { get; set; }
        public string type { get; set; }


        public ReportOutOfServiceInfo reportMonitoringInfo = new ReportOutOfServiceInfo();
        public List<Incident> ActualIncidents = new List<Incident>();
        public List<string> GroupIds;
        public List<AtmInfo> AtmInfoLst = new List<AtmInfo>();
        public List<AtmCountsGet.AtmCountsData> AtmCountsLst;
        public Dictionary<string, AtmGroupsGet.AtmGroup> groups;
        public AtmGroupsGet.Info groupTree;

        public DictionariesGet.Info dictionariesInfo;

        private Dictionary<string, string> failuresDict = new Dictionary<string, string>();

        public void CreateFailuresDict()
        {
            foreach (DictionaryGetDevicesTypes.Data data in dictionariesInfo.devicesTypes.data)
            {
                failuresDict.Add(data.name, data.id);
            }
        }

        List<ReportColums> reportColumns = new List<ReportColums>();

        public void Init(Report report)
        {
            reportMonitoringInfo.from = report.from;
            reportMonitoringInfo.to = report.to;
            reportMonitoringInfo.path = report.path;
            GroupIds = report.atmsGroupsId;
            reportMonitoringInfo.data.ActualIncidents = this.ActualIncidents;
            reportMonitoringInfo.data.AtmInfoLst = this.AtmInfoLst;
        }

        public Stylesheet MakeStyleSheet()
        {
            Stylesheet stylesheet = new Stylesheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" }
            };

            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts = new Fonts() { Count = 2U, KnownFonts = true };

            Font font11D = new Font();
            FontSize fontSize11D = new FontSize() { Val = 11D };

            font11D.Append(fontSize11D);

            Font font11DBold = new Font();
            Bold bold11DBold = new Bold();
            FontSize fontSize11DBold = new FontSize() { Val = 11D };

            font11DBold.Append(bold11DBold);
            font11DBold.Append(fontSize11DBold);

            fonts.Append(font11D);
            fonts.Append(font11DBold);

            Fills fills = new Fills() { Count = 1U };

            Fill fill = new Fill();
            PatternFill patternFill = new PatternFill() { PatternType = PatternValues.None };

            fills.Append(fill);

            Borders borders = new Borders() { Count = (UInt32Value)1U };

            Border border = new Border();
            LeftBorder leftBorder = new LeftBorder() { Style = BorderStyleValues.Dotted };
            RightBorder rightBorder = new RightBorder() { Style = BorderStyleValues.Dotted };
            TopBorder topBorder = new TopBorder() { Style = BorderStyleValues.Dotted };
            BottomBorder bottomBorder = new BottomBorder() { Style = BorderStyleValues.Dotted };
            DiagonalBorder diagonalBorder = new DiagonalBorder() { Style = BorderStyleValues.Dotted };


            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);

            borders.Append(border);

            border = new Border();
            leftBorder = new LeftBorder() { Style = BorderStyleValues.Medium };
            rightBorder = new RightBorder() { Style = BorderStyleValues.Medium };
            topBorder = new TopBorder() { Style = BorderStyleValues.Medium };
            bottomBorder = new BottomBorder() { Style = BorderStyleValues.Medium };
            diagonalBorder = new DiagonalBorder() { Style = BorderStyleValues.Medium };

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);
            borders.Append(border);

            border = new Border();
            leftBorder = new LeftBorder() { Style = BorderStyleValues.Thin };
            rightBorder = new RightBorder() { Style = BorderStyleValues.Medium };
            topBorder = new TopBorder() { Style = BorderStyleValues.Thin };
            bottomBorder = new BottomBorder() { Style = BorderStyleValues.Thin };
            diagonalBorder = new DiagonalBorder() { Style = BorderStyleValues.Thin };

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);
            borders.Append(border);

            border = new Border();
            leftBorder = new LeftBorder();
            rightBorder = new RightBorder();
            topBorder = new TopBorder() { Style = BorderStyleValues.Medium };
            bottomBorder = new BottomBorder();
            diagonalBorder = new DiagonalBorder();

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);
            borders.Append(border);

            border = new Border();
            leftBorder = new LeftBorder() { Style = BorderStyleValues.Thin };
            rightBorder = new RightBorder() { Style = BorderStyleValues.Thin };
            topBorder = new TopBorder() { Style = BorderStyleValues.Thin };
            bottomBorder = new BottomBorder() { Style = BorderStyleValues.Thin };
            diagonalBorder = new DiagonalBorder() { Style = BorderStyleValues.Thin };


            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);

            borders.Append(border);


            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)4U };

            // 0U - Default format.
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFont = true };

            // 1U - Default format with center-center alignment.
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, ApplyFont = true, ApplyAlignment = true };
            cellFormat2.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center });

            // 2U - Default format with center-top alignment.
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            cellFormat3.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top });

            // 3U - Bold font format.
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFont = true };

            // 4U - Bold font format with center alignment.
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, ApplyFont = true, ApplyAlignment = true };
            cellFormat5.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true });

            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, ApplyFont = true, ApplyAlignment = true };
            cellFormat6.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true });

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, ApplyFont = true, ApplyAlignment = true };
            cellFormat7.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center });


            cellFormats.Append(cellFormat1);
            cellFormats.Append(cellFormat2);
            cellFormats.Append(cellFormat3);
            cellFormats.Append(cellFormat4);
            cellFormats.Append(cellFormat5);
            cellFormats.Append(cellFormat6);
            cellFormats.Append(cellFormat7);

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellFormats);

            return stylesheet;
        }

        public bool MakeAnExcel()
        {
            string[] fromArray = reportMonitoringInfo.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = reportMonitoringInfo.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            reportMonitoringInfo.path = reportMonitoringInfo.path.Replace("/", "\\") + "\\OUS_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(reportMonitoringInfo.path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart;
                WorksheetPart worksheetPart;
                WorkbookStylesPart workbookStylesPart;

                workbookpart = spreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                workbookStylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = MakeStyleSheet();
                Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());


                Sheet sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    Name = "Служба мониторинга",
                    SheetId = (uint)1
                };
                sheets.Append(sheet);

                var gr = groupTree.usersGroup[0].atmGroups.Where(grr => GroupIds.Contains(grr.id.ToString())).ToList();
                var headGroups = groups.Where(gr1 => gr.Contains(gr1.Value)).ToList();

                CreateHeaderRow(worksheetPart);
                CreateFooterRow(worksheetPart);
                CreateTotalRows(worksheetPart, headGroups);
                CreateFooterRow(worksheetPart);

               
                foreach (var group in headGroups)
                {                    
                    CreateNADRows(worksheetPart, group);
                    CreateNAARows(worksheetPart, group);
                    CreateCommRows(worksheetPart, group);
                    CreateEncashRows(worksheetPart, group);
                    CreateFooterRow(worksheetPart);
                }


                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 1, 12);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 2, 42);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 3, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 4, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 5, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 6, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 7, 34);
                ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 8, 34);
                workbookpart.Workbook.Save();
            }

            reportMonitoringInfo.data.ActualIncidents.Clear();
            reportMonitoringInfo.data.AtmInfoLst.Clear();

            return true;
        }

        public bool SearchAtmId(string Id,AtmGroupsGet.AtmGroup group)
        {
            var result = false;
            if (group.atmIds.Contains(Convert.ToInt32(Id)))
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

            string title = "Отчет о простаивающих банкоматах";
            for (int i = 1; i <= 7; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A1", "H1");
        }

        private void CreateCommRows(WorksheetPart worksheetPart, KeyValuePair<string, AtmGroupsGet.AtmGroup> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            AtmInfo Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #region title
            string title = "Отсутствие связи";
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #endregion


            var incidents = ActualIncidents.Where(incident => incident.deviceTypeId == failuresDict["Communications"] && SearchAtmId(incident.atmId,group.Value));

            #region head
            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.CreateCell(row, 1, row.RowIndex, "№ATM", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, "Aдрес банкомата", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, "Место расположения", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, "Описание неисправности", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, "Модель", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, "Дата", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 7, row.RowIndex, "Предельный срок восстановления", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, "Текущее состояние", CellValues.String, 4U);
            }
            else
            {
                title = "Нет";
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                return;
            }

            #endregion

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = AtmInfoLst.Where(atm => atm.id == incident.atmId).First();

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.id, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.geoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, "нет связи", CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);
                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.recoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd hh:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);

            }

        }

        private void CreateNADRows(WorksheetPart worksheetPart, KeyValuePair<string, AtmGroupsGet.AtmGroup> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            AtmInfo Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #region title
            string title = group.Value.name;
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            title = "Ремонт.Недоступные на выдачу";
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            #endregion

            var incidents = ActualIncidents.Where(
                incident => (incident.deviceTypeId == failuresDict["CardReader"] ||
                                                      (incident.deviceTypeId == failuresDict["JournalPrinter"] && !(incident.subject == "Ж.Принтер: Мало бумаги -> FLM ж.принтер" || incident.subject == "Ж.Принтер: Бумага закончилась -> FLM ж.принтер"))) &&
                                                      SearchAtmId(incident.atmId, group.Value)
                                                      );

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.CreateCell(row, 1, row.RowIndex, "№ATM", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, "Aдрес банкомата", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, "Место расположения", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, "Описание неисправности", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, "Модель", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, "Дата", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 7, row.RowIndex, "Предельный срок восстановления", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, "Текущее состояние", CellValues.String, 4U);
            }
            else
            {
                title = "Нет";
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = AtmInfoLst.Where(atm => atm.id == incident.atmId).First();

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.id, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.geoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, incident.subject, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.recoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

        }

        private void CreateNAARows(WorksheetPart worksheetPart, KeyValuePair<string, AtmGroupsGet.AtmGroup> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            AtmInfo Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #region title
            var title = "Ремонт.Не доступныe на прием";
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            #endregion

            var incidents = ActualIncidents.Where(
                incident => (incident.deviceTypeId == failuresDict["BNA"] ||
                                                      (incident.deviceTypeId == failuresDict["ReceiptPrinter"] && !(incident.subject == "Ч.Принтер: Бумага закончилась -> FLM ч.принтер" || incident.subject == "Ч.Принтер: Мало бумаги -> FLM ч.принтер"))) &&
                                                      SearchAtmId(incident.atmId, group.Value)
                                                      );

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.CreateCell(row, 1, row.RowIndex, "№ATM", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, "Aдрес банкомата", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, "Место расположения", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, "Описание неисправности", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, "Модель", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, "Дата", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 7, row.RowIndex, "Предельный срок восстановления", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, "Текущее состояние", CellValues.String, 4U);
            }
            else
            {
                title = "Нет";
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = AtmInfoLst.Where(atm => atm.id == incident.atmId).First();

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.id, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.geoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, incident.subject, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.recoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

        }

        private void CreateEncashRows(WorksheetPart worksheetPart, KeyValuePair<string, AtmGroupsGet.AtmGroup> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            AtmInfo Atm;
            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            #region title
            var title = "Инкассация";
            for (int i = 1; i <= 8; i++)
            {
                if (i > 1)
                    title = "";
                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());

            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            #endregion
            var role = dictionariesInfo.userRoles.data.Where(typ => typ.description == "Инкассация").First();
            var incidents = ActualIncidents.Where(incident => (incident.userRoleId == role.id) && SearchAtmId(incident.atmId, group.Value));

            if (!incidents.IsNullOrEmpty())
            {
                ExcelHelper.CreateCell(row, 1, row.RowIndex, "№ATM", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, "Aдрес банкомата", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, "Место расположения", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, "Описание неисправности", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, "Модель", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, "Дата", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 7, row.RowIndex, "Предельный срок восстановления", CellValues.String, 4U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, "Текущее состояние", CellValues.String, 4U);
            }
            else
            {
                title = "Нет";
                for (int i = 1; i <= 8; i++)
                {
                    if (i > 1)
                        title = "";
                    ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
                }
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex.ToString(), "H" + row.RowIndex.ToString());
                return;
            }

            foreach (Incident incident in incidents)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                Atm = AtmInfoLst.Where(atm => atm.id == incident.atmId).First();

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.id, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.geoAddress, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, Atm.place, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, incident.subject, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, Atm.model, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 6, row.RowIndex, incident.timeCreated, CellValues.String, 5U);

                var date = DateTime.Parse(incident.timeCreated);
                var hours = 0.0;
                if (double.TryParse(Atm.recoveryTime, out hours))
                    date.AddHours(hours);

                ExcelHelper.CreateCell(row, 7, row.RowIndex, date.ToString("yyyy-MM-dd HH:mm:ss"), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, incident.comments, CellValues.String, 5U);
            }

        }

        private void CreateTotalRows(WorksheetPart worksheetPart,List<KeyValuePair<string, AtmGroupsGet.AtmGroup>> group)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
           
            row = (Row)sheetData.LastChild;
            sheetData.Append(new Row() { RowIndex = row.RowIndex + 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
           

            ExcelHelper.CreateCell(row, 1, row.RowIndex, "", CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 2, row.RowIndex, "Сервис (Не работают на выдачу /работают на выдачу, но не работают на прием)", CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 3, row.RowIndex, "Связь", CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 4, row.RowIndex, "Инкассация", CellValues.String, 4U);
            

            foreach (var grp  in group)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                var incAcc = ActualIncidents.Where(
               incident => (incident.deviceTypeId == failuresDict["BNA"] ||
                                                     (incident.deviceTypeId == failuresDict["ReceiptPrinter"] && !(incident.subject == "Ч.Принтер: Бумага закончилась -> FLM ч.принтер" || incident.subject == "Ч.Принтер: Мало бумаги -> FLM ч.принтер"))) &&
                                                     SearchAtmId(incident.atmId, grp.Value)
                                                     ).ToList();
                var incDisp = ActualIncidents.Where(
                incident => (incident.deviceTypeId == failuresDict["CardReader"] ||
                                                      (incident.deviceTypeId == failuresDict["JournalPrinter"] && !(incident.subject == "Ж.Принтер: Мало бумаги -> FLM ж.принтер" || incident.subject == "Ж.Принтер: Бумага закончилась -> FLM ж.принтер"))) &&
                                                       SearchAtmId(incident.atmId, grp.Value)
                                                      ).ToList();

                var comms = ActualIncidents.Where(incident => incident.deviceTypeId == failuresDict["Communications"] && SearchAtmId(incident.atmId, grp.Value)).ToList();

                var role = dictionariesInfo.userRoles.data.Where(typ => typ.description == "Инкассация").First();
                var encash = ActualIncidents.Where(incident => incident.userRoleId == role.id && SearchAtmId(incident.atmId, grp.Value)).ToList();


                ExcelHelper.CreateCell(row, 1, row.RowIndex, grp.Value.name, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, incDisp.Count.ToString() + "/" + incAcc.Count.ToString(), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 3, row.RowIndex, comms.Count.ToString(), CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, encash.Count.ToString(), CellValues.String, 5U);                
            }

        }

        private void CreateFooterRow(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            row = (Row)sheetData.LastChild;
            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <=8; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }
        }
    }
}