using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

using M3Atms;
using M3Utils;

namespace M3Reports
{
    public class ReportCounts : ReportBuilder
    {
        public List<ReportColumns> reportColumns;
        public List<ReportColumns> reportBNAColumns;

        internal override void MakeAnExcel()
        {
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\CNT_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

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
                workbookStylesPart.Stylesheet = ExcelHelper.MakeStyleSheet();

                Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                Sheet sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    Name = ReportsSource.Dispenser,
                    SheetId = 1
                };

                sheets.Append(sheet);

                this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportCountsColumn.xml");
                this.reportBNAColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportBNACountsColumn.xml");

                this.CreateHeaderRow(worksheetPart);
                this.CreateDataRows(worksheetPart);

                for (int i = 1; i <= this.reportColumns.Count(); i++)
                    ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportColumns[i - 1].width);

                worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    Name = "BNA",
                    SheetId = 2
                };

                sheets.Append(sheet);

                this.CreateBNAHeaderRow(worksheetPart);
                this.CreateBNADataRows(worksheetPart);
                for (int i = 1; i <= this.reportBNAColumns.Count(); i++)
                    ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportBNAColumns[i - 1].width);

                workbookpart.Workbook.Save();
            }
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = ReportsSource.ReportOnCashBalances_Dispenser;

            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                if (i > 1)
                    title = "";

                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }

            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportColumns[0].localtion + row.RowIndex, this.reportColumns[this.reportColumns.Count - 1].localtion + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            ExcelHelper.CreateCell(row, 1, row.RowIndex, ReportsSource.ATM, CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 2, row.RowIndex, ReportsSource.Address, CellValues.String, 4U);

            ExcelHelper.CreateCell(row, 3, row.RowIndex, ReportsSource.Cassette + " 1", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "C2", "F2");
            ExcelHelper.CreateCell(row, 7, row.RowIndex, ReportsSource.Cassette + " 2", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "G2", "J2");
            ExcelHelper.CreateCell(row, 11, row.RowIndex, ReportsSource.Cassette + " 3", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "K2", "N2");
            ExcelHelper.CreateCell(row, 15, row.RowIndex, ReportsSource.Cassette + " 4", CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 18, row.RowIndex, "", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "O2", "R2");

            ExcelHelper.CreateCell(row, 19, row.RowIndex, ReportsSource.TotalAmount, CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 20, row.RowIndex, "", CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 21, row.RowIndex, "", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "S2", "U2");

            ExcelHelper.CreateCell(row, 22, row.RowIndex, ReportsSource.CashBalance, CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 23, row.RowIndex, "", CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 24, row.RowIndex, "", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "V2", "X2");

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportColumns[i - 1].title, CellValues.String, 4U);
            }
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A2", "A3");
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "B2", "B3");

        }

        private void CreateBNAHeaderRow(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = ReportsSource.ReportOnCashBalancesBNA;

            for (int i = 1; i <= this.reportBNAColumns.Count; i++)
            {
                if (i > 1)
                    title = "";

                ExcelHelper.CreateCell(row, i, row.RowIndex, title, CellValues.String, 4U);
            }

            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportBNAColumns[0].localtion + row.RowIndex, this.reportBNAColumns[this.reportBNAColumns.Count - 1].localtion + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            ExcelHelper.CreateCell(row, 1, row.RowIndex, ReportsSource.ATM, CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 2, row.RowIndex, ReportsSource.Address, CellValues.String, 4U);

            ExcelHelper.CreateCell(row, 3, row.RowIndex, "", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "C2", "E2");

            ExcelHelper.CreateCell(row, 6, row.RowIndex, "RUB", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "F2", "L2");
            ExcelHelper.CreateCell(row, 13, row.RowIndex, "USD", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "M2", "T2");
            ExcelHelper.CreateCell(row, 21, row.RowIndex, "EUR", CellValues.String, 4U);
            //   M3Utils.ExcelHelper.CreateCell(row, 25, row.RowIndex, "", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "U2", "AB2");

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportBNAColumns.Count; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, this.reportBNAColumns[i - 1].title, CellValues.String, 4U);
            }

            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A2", "A3");
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "B2", "B3");
        }

        private void CreateDataRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            int count;
            int load;
            int nominal;
            int sum;
            int sum_load;
            int RUBsum = 0;
            int EURsum = 0;
            int USDsum = 0;
            int RUBload = 0;
            int EURload = 0;
            int USDload = 0;
            int totalRUBsum = 0;
            int totalEURsum = 0;
            int totalUSDsum = 0;
            int totalRUBload = 0;
            int totalEURload = 0;
            int totalUSDload = 0;

            foreach (CountsGet.AtmCountsData data in this.Data.AtmCounts)
            {
                RUBsum = 0;
                USDsum = 0;
                EURsum = 0;
                RUBload = 0;
                EURload = 0;
                USDload = 0;

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                Atm = this.Data.AtmInfo.First(atm => atm.Id == data.atmId);

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.Vizname, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);

                ExcelHelper.CreateCell(row, 3, row.RowIndex, data.cassete_1_Currency, CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 4, row.RowIndex, data.сassete_1_Value, CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 5, row.RowIndex, data.total_RemainCass_Pos1, CellValues.String, 1U);
                if (Int32.TryParse(data.total_RemainCass_Pos1, out count) && Int32.TryParse(data.total_LoadCass_Pos1, out load) && Int32.TryParse(data.сassete_1_Value, out nominal))
                {
                    sum = count * nominal;
                    sum_load = load * nominal;
                    switch (data.cassete_1_Currency)
                    {
                        case "RUB":
                            RUBsum += sum;
                            RUBload += sum_load;
                            break;
                        case "USD":
                            USDsum += sum;
                            USDload += sum_load;
                            break;
                        case "EUR":
                            EURsum += sum;
                            EURload += sum_load;
                            break;
                    }
                    ExcelHelper.CreateCell(row, 6, row.RowIndex, sum.ToString(), CellValues.String, 5U);
                }

                ExcelHelper.CreateCell(row, 7, row.RowIndex, data.cassete_2_Currency, CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 8, row.RowIndex, data.сassete_2_Value, CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 9, row.RowIndex, data.total_RemainCass_Pos2, CellValues.String, 1U);
                if (Int32.TryParse(data.total_RemainCass_Pos2, out count) && Int32.TryParse(data.total_LoadCass_Pos2, out load) && Int32.TryParse(data.сassete_2_Value, out nominal))
                {
                    sum = count * nominal;
                    sum_load = load * nominal;
                    switch (data.cassete_2_Currency)
                    {
                        case "RUB":
                            RUBsum += sum;
                            RUBload += sum_load;
                            break;
                        case "USD":
                            USDsum += sum;
                            USDload += sum_load;
                            break;
                        case "EUR":
                            EURsum += sum;
                            EURload += sum_load;
                            break;
                    }
                    ExcelHelper.CreateCell(row, 10, row.RowIndex, sum.ToString(), CellValues.String, 5U);
                }

                ExcelHelper.CreateCell(row, 11, row.RowIndex, data.cassete_3_Currency, CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 12, row.RowIndex, data.сassete_3_Value, CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 13, row.RowIndex, data.total_RemainCass_Pos3, CellValues.String, 1U);
                if (Int32.TryParse(data.total_RemainCass_Pos3, out count) && Int32.TryParse(data.total_LoadCass_Pos3, out load) && Int32.TryParse(data.сassete_3_Value, out nominal))
                {
                    sum = count * nominal;
                    sum_load = load * nominal;
                    switch (data.cassete_3_Currency)
                    {
                        case "RUB":
                            RUBsum += sum;
                            RUBload += sum_load;
                            break;
                        case "USD":
                            USDsum += sum;
                            USDload += sum_load;
                            break;
                        case "EUR":
                            EURsum += sum;
                            EURload += sum_load;
                            break;
                    }
                    ExcelHelper.CreateCell(row, 14, row.RowIndex, sum.ToString(), CellValues.String, 5U);
                }
                ExcelHelper.CreateCell(row, 15, row.RowIndex, data.cassete_4_Currency, CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 16, row.RowIndex, data.сassete_4_Value, CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 17, row.RowIndex, data.total_RemainCass_Pos4, CellValues.String, 1U);
                if (Int32.TryParse(data.total_RemainCass_Pos4, out count) && Int32.TryParse(data.total_LoadCass_Pos4, out load) && Int32.TryParse(data.сassete_4_Value, out nominal))
                {
                    sum = count * nominal;
                    sum_load = load * nominal;
                    switch (data.cassete_4_Currency)
                    {
                        case "RUB":
                            RUBsum += sum;
                            RUBload += sum_load;
                            break;
                        case "USD":
                            USDsum += sum;
                            USDload += sum_load;
                            break;
                        case "EUR":
                            EURsum += sum;
                            EURload += sum_load;
                            break;
                    }
                    ExcelHelper.CreateCell(row, 18, row.RowIndex, (count * nominal).ToString(), CellValues.String, 5U);
                }

                ExcelHelper.CreateCell(row, 19, row.RowIndex, RUBsum.ToString(), CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 20, row.RowIndex, USDsum.ToString(), CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 21, row.RowIndex, EURsum.ToString(), CellValues.String, 5U);

                ExcelHelper.CreateCell(row, 22, row.RowIndex, ((RUBload != 0) ? ((double)RUBsum / (double)RUBload) : 0).ToString("p1"), CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 23, row.RowIndex, ((USDload != 0) ? ((double)USDsum / (double)USDload) : 0).ToString("p1"), CellValues.String, 1U);
                ExcelHelper.CreateCell(row, 24, row.RowIndex, ((EURload != 0) ? ((double)EURsum / (double)EURload) : 0).ToString("p1"), CellValues.String, 5U);

                totalRUBsum += RUBsum;
                totalEURsum += EURsum;
                totalUSDsum += USDsum;
                totalRUBload += RUBload;
                totalEURload += EURload;
                totalUSDload += USDload;
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;

            ExcelHelper.CreateCell(row, 1, row.RowIndex, ReportsSource.TotalAmountOfBalancesForAGroupOfAtms, CellValues.String, 4U);
            ExcelHelper.CreateCell(row, 2, row.RowIndex, "", CellValues.String, 4U);
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "B" + row.RowIndex);

            for (int i = 3; i < 19; i++)
            {
                if ((i - 2) % 4 == 0)
                    ExcelHelper.CreateCell(row, i, row.RowIndex, "-", CellValues.String, 5U);
                else
                    ExcelHelper.CreateCell(row, i, row.RowIndex, "-", CellValues.String, 1U);
            }

            ExcelHelper.CreateCell(row, 19, row.RowIndex, totalRUBsum.ToString(), CellValues.String, 1U);
            ExcelHelper.CreateCell(row, 20, row.RowIndex, totalUSDsum.ToString(), CellValues.String, 1U);
            ExcelHelper.CreateCell(row, 21, row.RowIndex, totalEURsum.ToString(), CellValues.String, 5U);

            ExcelHelper.CreateCell(row, 22, row.RowIndex, ((RUBload != 0) ? ((double)totalRUBsum / (double)totalRUBload) : 0).ToString("p1"), CellValues.String, 1U);
            ExcelHelper.CreateCell(row, 23, row.RowIndex, ((USDload != 0) ? ((double)totalUSDsum / (double)totalUSDload) : 0).ToString("p1"), CellValues.String, 1U);
            ExcelHelper.CreateCell(row, 24, row.RowIndex, ((EURload != 0) ? ((double)totalEURsum / (double)totalEURload) : 0).ToString("p1"), CellValues.String, 5U);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }

        }

        private void CreateBNADataRows(WorksheetPart worksheetPart)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();
            Info Atm;
            row = (Row)sheetData.LastChild;

            int RUBsum = 0;
            int EURsum = 0;
            int USDsum = 0;

            //     public List<AtmInfo> AtmInfoLst;

            foreach (BNACountsGet.AtmBNACountsData data in this.Data.AtmBNACounts)
            {
                RUBsum = 0;
                USDsum = 0;
                EURsum = 0;

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                Atm = this.Data.AtmInfo.First(atm => atm.Id == data.id);

                ExcelHelper.CreateCell(row, 1, row.RowIndex, Atm.Vizname, CellValues.String, 5U);
                ExcelHelper.CreateCell(row, 2, row.RowIndex, Atm.GeoAddress, CellValues.String, 5U);

                int iLimit = 0;

                foreach (Info info in this.Data.AtmInfo)
                {
                    if (data.id == info.Id)
                    {
                        iLimit = Convert.ToInt32(info.Bnalimit);
                        break;
                    }
                }

                int iPaperCounter = 0;

                for (int i = 0; i < data.RUB.Count; i++)
                {
                    iPaperCounter += int.Parse(data.RUB[i].position_0_DepositCounts);
                    switch (data.RUB[i].currencyValue)
                    {
                        case 10:
                            ExcelHelper.CreateCell(row, 6, row.RowIndex, data.RUB[i].position_0_DepositCounts, CellValues.String, 1U);
                            RUBsum += 10 * int.Parse(data.RUB[i].position_0_DepositCounts);
                            break;

                        case 50:
                            ExcelHelper.CreateCell(row, 7, row.RowIndex, data.RUB[i].position_0_DepositCounts, CellValues.String, 1U);
                            RUBsum += 50 * int.Parse(data.RUB[i].position_0_DepositCounts);
                            break;
                        case 100:
                            ExcelHelper.CreateCell(row, 8, row.RowIndex, data.RUB[i].position_0_DepositCounts, CellValues.String, 1U);
                            RUBsum += 100 * int.Parse(data.RUB[i].position_0_DepositCounts);
                            break;
                        case 500:
                            ExcelHelper.CreateCell(row, 9, row.RowIndex, data.RUB[i].position_0_DepositCounts, CellValues.String, 1U);
                            RUBsum += 500 * int.Parse(data.RUB[i].position_0_DepositCounts);
                            break;
                        case 1000:
                            ExcelHelper.CreateCell(row, 10, row.RowIndex, data.RUB[i].position_0_DepositCounts, CellValues.String, 1U);
                            RUBsum += 1000 * int.Parse(data.RUB[i].position_0_DepositCounts);
                            break;
                        case 5000:
                            ExcelHelper.CreateCell(row, 11, row.RowIndex, data.RUB[i].position_0_DepositCounts, CellValues.String, 1U);
                            RUBsum += 5000 * int.Parse(data.RUB[i].position_0_DepositCounts);
                            break;
                    }
                }

                ExcelHelper.CreateCell(row, 12, row.RowIndex, RUBsum.ToString(), CellValues.String, 5U);

                for (int i = 0; i < data.USD.Count; i++)
                {
                    iPaperCounter += int.Parse(data.USD[i].position_0_DepositCounts);
                    switch (data.USD[i].currencyValue)
                    {
                        case 1:
                            ExcelHelper.CreateCell(row, 13, row.RowIndex, data.USD[i].position_0_DepositCounts, CellValues.String, 1U);
                            USDsum += int.Parse(data.USD[i].position_0_DepositCounts);
                            break;
                        case 2:
                            ExcelHelper.CreateCell(row, 14, row.RowIndex, data.USD[i].position_0_DepositCounts, CellValues.String, 1U);
                            USDsum += 2 * int.Parse(data.USD[i].position_0_DepositCounts);
                            break;
                        case 5:
                            ExcelHelper.CreateCell(row, 15, row.RowIndex, data.USD[i].position_0_DepositCounts, CellValues.String, 1U);
                            USDsum += 5 * int.Parse(data.USD[i].position_0_DepositCounts);
                            break;
                        case 10:
                            ExcelHelper.CreateCell(row, 16, row.RowIndex, data.USD[i].position_0_DepositCounts, CellValues.String, 1U);
                            USDsum += 10 * int.Parse(data.USD[i].position_0_DepositCounts);
                            break;
                        case 20:
                            ExcelHelper.CreateCell(row, 17, row.RowIndex, data.USD[i].position_0_DepositCounts, CellValues.String, 1U);
                            USDsum += 20 * int.Parse(data.USD[i].position_0_DepositCounts);
                            break;
                        case 50:
                            ExcelHelper.CreateCell(row, 18, row.RowIndex, data.USD[i].position_0_DepositCounts, CellValues.String, 1U);
                            USDsum += 50 * int.Parse(data.USD[i].position_0_DepositCounts);
                            break;
                        case 100:
                            ExcelHelper.CreateCell(row, 19, row.RowIndex, data.USD[i].position_0_DepositCounts, CellValues.String, 1U);
                            USDsum += 100 * int.Parse(data.USD[i].position_0_DepositCounts);
                            break;
                    }
                }
                ExcelHelper.CreateCell(row, 20, row.RowIndex, USDsum.ToString(), CellValues.String, 5U);

                for (int i = 0; i < data.EUR.Count; i++)
                {
                    iPaperCounter += int.Parse(data.EUR[i].position_0_DepositCounts);
                    switch (data.EUR[i].currencyValue)
                    {
                        case 5:
                            ExcelHelper.CreateCell(row, 21, row.RowIndex, data.EUR[i].position_0_DepositCounts, CellValues.String, 1U);
                            EURsum += 5 * int.Parse(data.EUR[i].position_0_DepositCounts);
                            break;
                        case 10:
                            ExcelHelper.CreateCell(row, 22, row.RowIndex, data.EUR[i].position_0_DepositCounts, CellValues.String, 1U);
                            EURsum += 10 * int.Parse(data.EUR[i].position_0_DepositCounts);
                            break;
                        case 20:
                            ExcelHelper.CreateCell(row, 23, row.RowIndex, data.EUR[i].position_0_DepositCounts, CellValues.String, 1U);
                            EURsum += 20 * int.Parse(data.EUR[i].position_0_DepositCounts);
                            break;
                        case 50:
                            ExcelHelper.CreateCell(row, 24, row.RowIndex, data.EUR[i].position_0_DepositCounts, CellValues.String, 1U);
                            EURsum += 50 * int.Parse(data.EUR[i].position_0_DepositCounts);
                            break;
                        case 100:
                            ExcelHelper.CreateCell(row, 25, row.RowIndex, data.EUR[i].position_0_DepositCounts, CellValues.String, 1U);
                            EURsum += 100 * int.Parse(data.EUR[i].position_0_DepositCounts);
                            break;
                        case 200:
                            ExcelHelper.CreateCell(row, 26, row.RowIndex, data.EUR[i].position_0_DepositCounts, CellValues.String, 1U);
                            EURsum += 200 * int.Parse(data.EUR[i].position_0_DepositCounts);
                            break;
                        case 500:
                            ExcelHelper.CreateCell(row, 27, row.RowIndex, data.EUR[i].position_0_DepositCounts, CellValues.String, 1U);
                            EURsum += 500 * int.Parse(data.EUR[i].position_0_DepositCounts);
                            break;
                    }
                    ExcelHelper.CreateCell(row, 28, row.RowIndex, EURsum.ToString(), CellValues.String, 5U);
                }

                ExcelHelper.CreateCell(row, 3, row.RowIndex, (iPaperCounter).ToString(), CellValues.String, 2U);
                if (iLimit > 0)
                    ExcelHelper.CreateCell(row, 4, row.RowIndex, (((decimal)iPaperCounter / (decimal)iLimit)).ToString("p0"), CellValues.String, 2U);
                else
                    ExcelHelper.CreateCell(row, 4, row.RowIndex, "NoN", CellValues.String, 2U);

                ExcelHelper.CreateCell(row, 5, row.RowIndex, iLimit.ToString(), CellValues.String, 2U);

            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= this.reportBNAColumns.Count; i++)
            {
                ExcelHelper.CreateCell(row, i, row.RowIndex, "", CellValues.String, 6U);
            }
        }
    }
}