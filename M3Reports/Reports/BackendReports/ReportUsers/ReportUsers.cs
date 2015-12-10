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
    public class ReportUsers : ReportBuilder
    {
        public List<UserActionItem> userActionItems;
        public List<FuncDesctiptionItem> funcDesctiptionItems;
        public List<UserDescriptionItem> userDescriptionItems;

        private List<ReportColumns> reportColumns;

        public void ParseUserActions(XmlNode messageNode)
        {
            XmlNodeList userActionsXmlList = messageNode.SelectNodes("Request/Item");

            UserActionItem userAction;

            if (userActionsXmlList != null)
            {
                this.userActionItems = new List<UserActionItem>();

                for (int i = 0; i < userActionsXmlList.Count; i++)
                {
                    userAction = new UserActionItem()
                    {
                        UserId = userActionsXmlList[i].SelectSingleNode("UserID").InnerText,
                        FuncId = userActionsXmlList[i].SelectSingleNode("FuncID").InnerText,
                        IPAddr = userActionsXmlList[i].SelectSingleNode("IPAddr").InnerText,
                        DTIme = userActionsXmlList[i].SelectSingleNode("DateTime").InnerText,
                    };

                    this.userActionItems.Add(userAction);
                }
            }
        }

        public void ParseFuncDescriptions(XmlNode messageNode)
        {
            XmlNodeList funcDesctiptionXmlList = messageNode.SelectNodes("Request/Func");

            FuncDesctiptionItem funcDesctiptionItem;

            if (funcDesctiptionXmlList != null)
            {
                this.funcDesctiptionItems = new List<FuncDesctiptionItem>();

                for (int i = 0; i < funcDesctiptionXmlList.Count; i++)
                {
                    funcDesctiptionItem = new FuncDesctiptionItem()
                    {
                        Id = funcDesctiptionXmlList[i].SelectSingleNode("Id").InnerText,
                        Description = funcDesctiptionXmlList[i].SelectSingleNode("Descr").InnerText
                    };

                    this.funcDesctiptionItems.Add(funcDesctiptionItem);
                }
            }
        }

        public void ParseUserDescriptions(XmlNode messageNode)
        {
            XmlNodeList userDescriptionXmlList = messageNode.SelectNodes("Request/User");

            UserDescriptionItem userDescriptionItem;

            if (userDescriptionXmlList != null)
            {
                this.userDescriptionItems = new List<UserDescriptionItem>();

                for (int i = 0; i < userDescriptionXmlList.Count; i++)
                {
                    userDescriptionItem = new UserDescriptionItem()
                    {
                        Id = userDescriptionXmlList[i].SelectSingleNode("Id").InnerText,
                        Login = userDescriptionXmlList[i].SelectSingleNode("Login").InnerText,
                        FName = userDescriptionXmlList[i].SelectSingleNode("FName").InnerText,
                        LName = userDescriptionXmlList[i].SelectSingleNode("LName").InnerText,
                        SName = userDescriptionXmlList[i].SelectSingleNode("SName").InnerText,
                        Phone = userDescriptionXmlList[i].SelectSingleNode("Phone").InnerText,
                        Email = userDescriptionXmlList[i].SelectSingleNode("Email").InnerText,
                    };

                    this.userDescriptionItems.Add(userDescriptionItem);
                }
            }
        }

        internal override void MakeAnExcel()
        {   
            string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            this.Info.path = this.Info.path.Replace("/", "\\") + "\\USR_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

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
                    Name = ReportsSource.UserActions,
                    SheetId = 1
                };

                sheets.Append(sheet);

                this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportUsersColumn.xml");

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
                                ReportsSource.ReportUserActions,
                                this.Info.userId,
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

            foreach (UserActionItem userAction in this.userActionItems)
            {
                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;

                List<FuncDesctiptionItem> funcDesctiptionItems = this.funcDesctiptionItems.Where(descr => descr.Id == userAction.FuncId).ToList();

                description = (funcDesctiptionItems.Count > 0) ? funcDesctiptionItems.First().Description : "";

                M3Utils.ExcelHelper.CreateCell(row, 1, row.RowIndex, description, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 2, row.RowIndex, userAction.DTIme, CellValues.String, 5U);
                M3Utils.ExcelHelper.CreateCell(row, 3, row.RowIndex, userAction.IPAddr, CellValues.String, 5U);
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