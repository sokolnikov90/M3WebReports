using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using M3Atms;
using M3Incidents;
using M3Dictionaries;
using M3TransactionListGenerator;

namespace M3Reports
{
    using M3IPClient;

    public class ReportTransactionsGet : ReportBuilder
    {
        public List<MsgData> msgList;
        public List<TransactionRepItem> trList;
        public List<Info> atmLst = new List<Info>();

//        public struct ReportTransactionData
//        {
//            public M3Atms.Info atmInfo;
//            public List<ReportTransactionItem> reportTransactionItem;
//        }

//        public struct ReportTransactionInfo
//        {
//            public int isError;
//            public string from;
//            public string to;
//            public string path;
//            public List<ReportTransactionData> data;
//        }
//
//        object Info
//        {
//            get
//            {
//                return this.reportTransactionInfo;
//            }
//        }
//
//        public ReportTransactionInfo reportTransactionInfo = new ReportTransactionInfo();

        private void MakeTrLst()
        {
            this.trList = new List<TransactionRepItem>();
            TransactionList tr = new TransactionList(AppDomain.CurrentDomain.BaseDirectory + @"bin/M3Reports/TrxRepCfg.xml");
            tr.MakeTransactionList(this.msgList, this.atmLst, ref this.trList);
        }

        public void ParseMessage(XmlNode messageNode)
        {
            this.Info.isError = 0;
            try
            {
                XDocument xmlDoc = XDocument.Parse(messageNode.InnerXml);

                this.msgList = (from msg in xmlDoc.Root.Elements("Item")
                           select new MsgData()
                           {
                               dateTime = msg.Element("DateTime").Value,
                               atmId = msg.Element("AtmID").Value,
                               body = msg.Element("Body").Value,
                               direction = msg.Element("Direction").Value,
                               guid = msg.Element("Guid").Value,
                               sec = msg.Element("SEQ").Value
                           }).OrderBy(msg => msg.dateTime).ToList();

                List<MsgData> SortedMsgList = new List<MsgData>();


                foreach (MsgData msg in this.msgList)
                {
                    if (SortedMsgList.Contains(msg))
                        continue;
                    var query = from Msg in this.msgList
                                where Msg.dateTime == msg.dateTime
                                orderby Msg.sec
                                select Msg;
                    SortedMsgList.AddRange(query);
                }

                this.msgList = SortedMsgList;

                for (int i = 0; i < this.msgList.Count - 1; i++)
                {
                    if (this.msgList[i].dateTime.Equals(this.msgList[i + 1].dateTime))
                    {
                        if (this.msgList[i].body.Contains("3232") && this.msgList[i + 1].body.Contains("341C"))
                        {
                            MsgData temp = this.msgList[i];
                            this.msgList[i] = this.msgList[i + 1];
                            this.msgList[i + 1] = temp;

                            string tempSec = this.msgList[i].sec;
                            this.msgList[i].sec = this.msgList[i + 1].sec;
                            this.msgList[i + 1].sec = tempSec;

                        }
                    }
                }
            }
            catch (Exception exception)
            {
                M3Utils.Log.Instance.Info("ReportGetTransaction() exeption: " + exception.Message);
                this.Info.isError = 1;
            }
        }

        private List<Field> ndcFields = new List<Field>();
        private List<Field> ddcFields = new List<Field>();
        private Dictionary<string, string> StateDescrs = new Dictionary<string, string>();
        private List<string> endTitles = new List<string>();

        public void ParseXML()
        {
            XDocument xmlDocument = XDocument.Load(AppDomain.CurrentDomain.BaseDirectory + @"bin/M3Reports/TrxRepCfg.xml");
            XElement transactions = xmlDocument.Root.Element("Transactions");
            XElement proto = transactions.Element("DDC");
            var columns = from column in proto.Elements("Transaction")
                          select column;

            List<string> names = new List<string>();
            foreach (XElement column in columns)
            {
                if (names.Contains(column.Element("Name").Value))
                    continue;
                names.Add(column.Element("Name").Value);

                this.endTitles.Add(column.Element("EndTitleIndex").Value);
                var flds = from field in column.Element("Fields").Elements("Field")
                           select new Field
                           {
                               title = field.Value,
                               description = field.Attribute("description").Value,
                               trType = column.Element("Name").Value
                           };

                this.ddcFields.AddRange(flds);
            }



            proto = transactions.Element("NDC");
            columns = from column in proto.Elements("Transaction")
                      select column;
            names.Clear();
            foreach (XElement column in columns)
            {
                if (names.Contains(column.Element("Name").Value))
                    continue;
                names.Add(column.Element("Name").Value);

                this.endTitles.Add(column.Element("EndTitleIndex").Value);
                var flds = from field in column.Element("Fields").Elements("Field")
                           select new Field
                           {
                               title = field.Value,
                               description = field.Attribute("description").Value,
                               trType = column.Element("Name").Value
                           };
                this.ndcFields.AddRange(flds);
            }

            XElement States = xmlDocument.Root.Element("States");
            var states = from st in States.Elements("State")
                         select st;

            foreach (XElement st in states)
            {
                this.StateDescrs.Add(st.Attribute("num").Value, st.Value);
            }


        }

        internal override void MakeAnExcel()
        {
            if (M3UserSession.BankName == "Absolut")
            {
                this.MakeAnExcelForAbsolute();
                return;
            }

            this.MakeTrLst();

            List<string> trNames = new List<string>();

            foreach (TransactionRepItem tr in this.trList)
            {
                if (!trNames.Contains(tr.Name))
                    trNames.Add(tr.Name);
            }


            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\TRN_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

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

                this.ParseXML();
                if (trNames.Count == 0)
                {
                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheet sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                        Name = "empty",
                        SheetId = (uint)(1)
                    };
                    sheets.Append(sheet);

                    this.CreateHeaderRow(worksheetPart, this.Info.from, this.Info.to, "C", new List<Field>());

                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 1, 25);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 2, 25);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 3, 20);
                }

                for (int j = 0; j < trNames.Count; j++)
                {
                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheet sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                        Name = trNames[j],
                        SheetId = (uint)(j + 1)
                    };

                    sheets.Append(sheet);

                    var trDDCFields = (from tr in this.ddcFields
                                       where tr.trType == trNames[j]
                                       select tr).ToList();

                    var trNDCFields = (from tr in this.ndcFields
                                       where tr.trType == trNames[j]
                                       select tr).ToList();

                    this.CreateHeaderRow(worksheetPart, this.Info.from, this.Info.to, this.endTitles[j], trDDCFields);
                    List<TransactionRepItem> trLs = (from n in this.trList
                                                     where n.Name == trNames[j]
                                                     select n).ToList();

                    foreach (Info atm in this.atmLst)
                    {
                        if (atm.TreeId == "1") this.CreateDataRows(worksheetPart, atm, trLs, this.endTitles[j], trDDCFields);
                        else this.CreateDataRows(worksheetPart, atm, trLs, this.endTitles[j], trNDCFields);
                    }

                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 1, 25);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 2, 25);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 3, 20);
                    int columnIndx = 4;
                    foreach (Field fld in trDDCFields)
                    {
                        M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, columnIndx, fld.title.Length + 4);
                        columnIndx++;
                    }
                }
                workbookpart.Workbook.Save();
            }
        }

        public void MakeAnExcelForAbsolute()
        {
            this.MakeTrLst();

            List<string> trNames = new List<string>();

            foreach (TransactionRepItem tr in this.trList)
            {
                if (!trNames.Contains(tr.Name))
                    trNames.Add(tr.Name);
            }


            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\TRN_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

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

                this.ParseXML();
                if (trNames.Count == 0)
                {
                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheet sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                        Name = "empty",
                        SheetId = (uint)(1)
                    };
                    sheets.Append(sheet);

                    this.CreateHeaderRowForAbsolute(worksheetPart, this.Info.from, this.Info.to, "");

                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 1, 25);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 2, 25);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 3, 20);
                }

                for (int j = 0; j < trNames.Count; j++)
                {
                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheet sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                        Name = trNames[j],
                        SheetId = (uint)(j + 1)
                    };

                    sheets.Append(sheet);

                    this.CreateHeaderRowForAbsolute(worksheetPart, this.Info.from, this.Info.to, trNames[j]);
                    List<TransactionRepItem> trLs = (from n in this.trList
                                                     where n.Name == trNames[j]
                                                     select n).ToList();

                    foreach (Info atm in this.atmLst)
                        if (atm.TreeId == "1") this.CreateDataRowsForAbsolute(worksheetPart, atm, trLs);
                        else this.CreateDataRowsForAbsolute(worksheetPart, atm, trLs);

                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 1, 25);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 2, 25);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 3, 30);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 4, 20);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 5, 20);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 6, 20);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 7, 20);
                    M3Utils.ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, 8, 20);
                    
                }

                workbookpart.Workbook.Save();
            }
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart, string from, string to, string endTitle, List<Field> flds)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = "Отчет по транзакциям за период c " + from + " по " + to;
            M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, title, CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, "", CellValues.String, 4U);
            int columnIndx = 4;

            foreach (Field fld in flds)
            {
                M3Utils.ExcelHelper.CreateCell(row, columnIndx, row.RowIndex, "", CellValues.String, 4U);
                columnIndx++;
            }

            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, endTitle + row.RowIndex);


        }

        private void CreateDataRows(WorksheetPart worksheetPart, Info atm, List<TransactionRepItem> trLstParam, string endTitle, List<Field> flds)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            row = (Row)sheetData.LastChild;
            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = "Банкомат: " + atm.Vizname + "; Адрес: " + atm.GeoAddress;

            M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, title, CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, "", CellValues.String, 4U);
            int columnIndx = 4;
            foreach (Field fld in flds)
            {
                M3Utils.ExcelHelper.CreateCell(row, columnIndx, row.RowIndex, "", CellValues.String, 4U);
                columnIndx++;
            }
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, endTitle + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;


            M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, "Время начала", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, "Время окончания", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, "Статус", CellValues.String, 4U);
            columnIndx = 4;
            foreach (Field fld in flds)
            {
                M3Utils.ExcelHelper.CreateCell(row, columnIndx, row.RowIndex, fld.title, CellValues.String, 4U);
                columnIndx++;
            }

            trLstParam = (from n in trLstParam
                          where n.AtmID == atm.Id
                          select n).ToList();
            foreach (TransactionRepItem tr in trLstParam)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, tr.DateTimeStart, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, tr.DateTimeFinish, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, tr.Status.ToString(), CellValues.String, 5U);

                columnIndx = 4;
                string val = "";
                foreach (Field fld in flds)
                {
                    try
                    {
                        val = tr.Fields[fld.description];
                        if (fld.description == "State")
                        {
                            try
                            {
                                val = this.StateDescrs[val];
                            }
                            catch { val = tr.Fields[fld.description]; }
                        }
                    }
                    catch { val = "00"; }
                    M3Utils.ExcelHelper.CreateCell(row, columnIndx, row.RowIndex, val, CellValues.String, 5U);
                    columnIndx++;
                }


            }
            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;



            M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, "", CellValues.String, 6U);
            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, "", CellValues.String, 6U);
            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, "", CellValues.String, 6U);
            columnIndx = 4;
            foreach (Field fld in flds)
            {
                M3Utils.ExcelHelper.CreateCell(row, columnIndx, row.RowIndex, "", CellValues.String, 6U);
                columnIndx++;
            }

        }

        private void CreateHeaderRowForAbsolute(WorksheetPart worksheetPart, string from, string to,string type)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string title = "Отчет по операции ("+ type + ") за период c " + from + " по " + to;
            M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, title, CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "H" + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            
            M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A" + row.RowIndex, "C" + row.RowIndex);
            
            M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, "Kоличество", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "D" + row.RowIndex, "F" + row.RowIndex);

            M3Utils.ExcelHelper.CreateCell(row,7, row.RowIndex, "Cумма RUB", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 8, row.RowIndex, "", CellValues.String, 4U);
            M3Utils.ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "G" + row.RowIndex, "H" + row.RowIndex);

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 30D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, "Банкомат", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, "Регион", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, "Адрес", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, "Успешные", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, "Неуспешные по вине оборудования", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, "Неуспешные  по вине клиента", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 7, row.RowIndex, "Успешные", CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 8, row.RowIndex, "Неуспешные", CellValues.String, 4U);
        }

        private void CreateDataRowsForAbsolute(WorksheetPart worksheetPart, Info atm, List<TransactionRepItem> trLstParam)
        {
            Row row;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            row = (Row)sheetData.LastChild;
            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            trLstParam = (from n in trLstParam
                          where n.AtmID == atm.Id
                          select n).ToList();

            int successTrx = 0;
            double successTrxSum = 0;
            int userFaultTrx = 0;
            double userFaultTrxSum = 0;
            int hwFaultTrx = 0;
            double hwFaultTrxSum = 0;

            foreach (TransactionRepItem tr in trLstParam)
            {
                try
                {
                    if (tr.Fields.ContainsKey("State"))
                    {
                        switch (this.StateDescrs[tr.Fields["State"]])
                        {
                            case "OK":
                                successTrx++;
                                successTrxSum += Convert.ToDouble(tr.Fields["Amount"]);
                                break;
                            case "UserFault":
                                userFaultTrx++;
                                userFaultTrxSum += Convert.ToDouble(tr.Fields["Amount"]);
                                break;
                            case "HardWareFault":
                                hwFaultTrx++;
                                hwFaultTrxSum += Convert.ToDouble(tr.Fields["Amount"]);
                                break;
                            case "UndefFault":
                                break;

                        }
                    }
                    else
                    {
                        hwFaultTrx++;
                        hwFaultTrxSum += Convert.ToDouble(tr.Fields["Amount"]);
                    }
                }
                catch
                {}
            }       
           
            M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, atm.Vizname, CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, atm.Region, CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, atm.GeoAddress, CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 4, row.RowIndex, successTrx.ToString(), CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 6, row.RowIndex, userFaultTrx.ToString(), CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 5, row.RowIndex, hwFaultTrx.ToString(), CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 7, row.RowIndex, successTrxSum.ToString(), CellValues.String, 4U);
            M3Utils.ExcelHelper.CreateCell(row, 8, row.RowIndex, (userFaultTrx+ hwFaultTrxSum).ToString(), CellValues.String, 4U);

        }

        private struct Field
        {
            public string title { get; set; }
            public string description { get; set; }
            public string trType { get; set; }
        }
    }
}