namespace M3Reports
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using M3Atms;

    using M3Dictionaries;

    using M3Incidents;

    using M3IPClient;

    using M3Utils;

    public enum NotchType : byte { begin = 0, end }

    public class Notch : IComparable
    {
        public NotchType type { get; set; }
        public DateTime time { get; set; }
        public bool deleted { get; set; }

        public int CompareTo(Object obj)
        {
            Notch notch = (Notch)obj;

            if (this.time > notch.time)
                return 1;
            if (this.time == notch.time)
            {
                if (this.type < notch.type)
                    return 1;
                if (this.type == notch.type)
                    return 0;
                if (this.type > notch.type)
                    return -1;                            
            }
            if (this.time < notch.time)
                return -1;

            throw new NotImplementedException();
        }
    }

    //Времена доступностей
    struct AvailsTime
    {
        public string AtmId { get; set; }

        public TimeSpan TechAvail { get; set; }
        public TimeSpan CashInAvail { get; set; }
        public TimeSpan DispenseAvail { get; set; }
        public TimeSpan PaymentAvail { get; set; }
        public TimeSpan StatAccountAvail { get; set; }
        public TimeSpan EncashAvail { get; set; }
        public TimeSpan CombiAvail { get; set; }
        public TimeSpan LockInterval { get; set; }
    }

    public struct Avail
    {
        public double TechAvail { get; set; }
        public double CashInAvail { get; set; }
        public double DispenseAvail { get; set; }
        public double PaymentAvail { get; set; }
        public double StatAccountAvail { get; set; }
        public double EncashAvail { get; set; }
        public double FunctionalAvail { get; set; }
        public double LockedFactor { get; set; }
    }

    public static class ExtensionMethodsUtils
    {
        public static void Cut(this List<Notch> notches)
        {
            var DeletingNotches = from notch in notches
                                  where notch.deleted == true
                                  select notch;

            foreach (Notch notch in DeletingNotches.ToList())
                notches.Remove(notch);
        }

        public static TimeSpan GetAvail(this List<Notch> notches)
        {
            TimeSpan time = new TimeSpan(0, 0, 0, 0, 0);
            Notch CurrentEnd = null;

            for (int i = 0; i < notches.Count; i++)
            {

                if (notches[i].type == NotchType.end)
                    CurrentEnd = notches[i];
                else
                {
                    time = time.Add(notches[i].time.Subtract(CurrentEnd.time));
                }
            }

            return time;
        }
    }

    public class ReportAvailabilities : ReportBuilder
    {
        public struct ReportAvailData
        {
            public Info atmInfo;
            public Avail avail;
        }

        public List<ReportAvailData> data = new List<ReportAvailData>();

        public List<GroupsGet.AtmGroup> AtmGroups { get; set; }

        public DateTime FromDate;
        public DateTime ToDate;

        public List<string> AtmsId
        {
            get
            {
                return this.Info.atmsId;
            }
        }
        public TimeSpan StatisticsPeriod;

        public List<ReportColumns> reportColumns { get; set; }

        public int hasBna;
        public int hasDispenser;
        public int hasBNAandDispenser;

        public string BankName { get; set; }

        #region DeviceTypes
        const string AtmWorkFailure = "1";
        const string AtmConnectedFailure = "2";
        const string CardReaderFailure = "3";
        const string EncriptorFailure = "4";
        const string DispencerFailure = "5";
        const string JournalPrinterFailure = "6";
        const string CheckPrinterFailure = "7";
        const string CasseteEmptyFailure = "8";
        const string BNAFailure = "9";
        List<string> Failures = new List<string>() { AtmWorkFailure, AtmConnectedFailure,CardReaderFailure,EncriptorFailure,
                                                          DispencerFailure,JournalPrinterFailure,CheckPrinterFailure,CasseteEmptyFailure,BNAFailure};

        List<string> StatusId = new List<string>() { "3", "4", "5" };

        private Dictionary<string, int> failuresDict;

        public void CreateFailuresDict()
        {
            this.failuresDict = new Dictionary<string, int>();
            foreach (GetDevicesTypes.Data data in this.Data.DictionariesGet.DevicesTypes)
            {
                this.failuresDict.Add(data.name, data.id);
            }
        }

        public string LockedId;
        public void GetLockedTypeId()
        {
            foreach (GetTypes.Data data in this.Data.DictionariesGet.Types)
            {
                if (data.text == "ATMLocked")
                {
                    this.LockedId = data.id.ToString();
                }
            }
        }

        #endregion

        private static bool beginWritten;
        private static int NotchCount;
        private static Notch CurrentEnd;
        private static DateTime DateCreate;
        private static DateTime DateClose;

        public void Init()
        {
            this.FromDate = DateTime.Parse(this.Info.from);
            this.ToDate = DateTime.Parse(this.Info.to);
            this.CreateFailuresDict();
            this.GetLockedTypeId();
            this.InitStatisticPeriod();
        }

        private void InitStatisticPeriod()
        {
            this.StatisticsPeriod = this.ToDate.Subtract(this.FromDate);

            if (this.StatisticsPeriod.Ticks == 0) this.StatisticsPeriod = new TimeSpan(1);
        }

        private void SearchNotches(List<Notch> notches, Incident incident)
        {
            beginWritten = false;
            NotchCount = notches.Count;
            DateCreate = DateTime.Parse(incident.timeCreated);
            DateClose = this.ToDate;

            try
            {

                if (incident.timeClosed != string.Empty)
                    DateClose = DateTime.Parse(incident.timeClosed);


                for (int i = 0; i < NotchCount; i++)
                {
                    notches[i].deleted = false;
                    if (notches[i].type == NotchType.end)
                        CurrentEnd = notches[i];
                    else
                    {
                        if (!beginWritten)
                        {
                            if (DateCreate > CurrentEnd.time && DateCreate < notches[i].time)
                            {
                                beginWritten = true;
                                notches.Add(new Notch() { time = DateCreate, type = NotchType.begin });
                            }
                            else
                            {
                                if (DateCreate <= CurrentEnd.time)
                                    CurrentEnd.deleted = true;
                                else
                                    continue;
                            }
                        }

                        if (incident.timeClosed != string.Empty)
                        {

                            if (DateClose > CurrentEnd.time && DateClose < notches[i].time)
                            {
                                notches.Add(new Notch() { time = DateClose, type = NotchType.end });
                                if (beginWritten && CurrentEnd.time > DateCreate)
                                    CurrentEnd.deleted = true;
                                break;
                            }
                            else
                            {
                                if (DateClose >= notches[i].time)
                                {
                                    notches[i].deleted = true;
                                    if (beginWritten && CurrentEnd.time > DateCreate)
                                        CurrentEnd.deleted = true;
                                }
                                else
                                    CurrentEnd.deleted = false;
                            }
                        }
                        else
                        {

                            notches[i].deleted = true;
                        }
                    }

                }

                notches.Sort();


                notches.Cut();
            }
            catch (Exception exp)
            {
                Log.Instance.Info(this + ".SearchNotches() exeption: " + exp.Message);
            }
        }

        internal override void MakeAnExcel()
        {
            try
            {
                string[] fromArray = this.Info.from.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                string[] toArray = this.Info.to.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

                this.Info.path = this.Info.path.Replace("/", "\\") + "\\AVL_HR_" + fromArray[0].Replace("-", "").Substring(2) + fromArray[1].Replace(":", "") + "_" + toArray[0].Replace("-", "").Substring(2) + toArray[1].Replace(":", "") + ".xlsx";

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
                        Name = "Availability",
                        SheetId = 1
                    };

                    sheets.Append(sheet);

                    this.reportColumns = ReportDataProvider.ParseXML(@"bin/M3Reports/ReportAvailabilityColumns.xml");

                    if (M3UserSession.BankName == "Absolute")
                    {
                        this.CreateHeaderRow(worksheetPart, this.Info.from, this.Info.to);
                        this.CreateDataRowsForAbsolute(worksheetPart, this.data);
                    }
                    else
                    {
                        this.CreateHeaderRow(worksheetPart, this.Info.from, this.Info.to);
                        this.CreateDataRows(worksheetPart, this.data);
                    }

                    for (int i = 1; i <= this.reportColumns.Count(); i++)
                        ExcelHelper.SetColumnWidth(worksheetPart.Worksheet, i, this.reportColumns[i - 1].width);

                    workbookpart.Workbook.Save();
                }

                this.data.Clear();
            }
            catch (Exception exp)
            {
                Log.Instance.Info(
                    this + ".MakeAnExcel(...) exception:",
                    exp.Message,
                    exp.Source,
                    exp.StackTrace);
            }
        }

        private void CreateHeaderRow(WorksheetPart worksheetPart, string from, string to)
        {
            try
            {
                Row row;
                SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

                sheetData.Append(new Row() { RowIndex = 1, Height = 30D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                string title = String.Join(" ", new[]
                                        {
                                            ReportsSource.ReportOnAvailability,
                                            ReportsSource.From,
                                            from,
                                            ReportsSource.To,
                                            to
                                        });

                for (int i = 1; i <= this.reportColumns.Count; i++)
                {
                    if (i > 1)
                        title = "";

                    string rowIndex = row.RowIndex;
                    Cell cell;
                    Cell previousCell;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex)))
                    {
                        cell = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex));
                    }
                    else
                    {
                        previousCell = null;

                        for (int i1 = i; i1 > 0; i1--)
                        {
                            previousCell = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex));

                            if (previousCell != null)
                                break;
                        }

                        cell = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(i) + rowIndex };

                        row.InsertAfter(cell, previousCell);
                    }

                    cell.DataType = CellValues.String;
                    cell.StyleIndex = 4U;
                    cell.CellValue = new CellValue(title);
                }

                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, this.reportColumns[0].localtion + row.RowIndex, this.reportColumns[this.reportColumns.Count - 1].localtion + row.RowIndex);
            }
            catch (Exception exp)
            {
                Log.Instance.Info(
                    this + ".CreateHeaderRow(...) exception:",
                    exp.Message,
                    exp.Source,
                    exp.StackTrace);
            }
        }

        private void CreateDataRows(WorksheetPart worksheetPart, List<ReportAvailData> Data)
        {
            try
            {
                Row row;
                Row AvrgRow;
                SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

                row = (Row)sheetData.LastChild;

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                string rowIndex4 = row.RowIndex;
                Cell cell4;
                Cell previousCell4;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex4)))
                {
                    cell4 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex4));
                }
                else
                {
                    previousCell4 = null;

                    for (int i4 = 1; i4 > 0; i4--)
                    {
                        previousCell4 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i4) + rowIndex4));

                        if (previousCell4 != null)
                            break;
                    }

                    cell4 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(1) + rowIndex4 };

                    row.InsertAfter(cell4, previousCell4);
                }

                cell4.DataType = CellValues.String;
                cell4.StyleIndex = 4U;
                cell4.CellValue = new CellValue("");
                string rowIndex5 = row.RowIndex;
                Cell cell5;
                Cell previousCell5;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex5)))
                {
                    cell5 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex5));
                }
                else
                {
                    previousCell5 = null;

                    for (int i5 = 2; i5 > 0; i5--)
                    {
                        previousCell5 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i5) + rowIndex5));

                        if (previousCell5 != null)
                            break;
                    }

                    cell5 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(2) + rowIndex5 };

                    row.InsertAfter(cell5, previousCell5);
                }

                cell5.DataType = CellValues.String;
                cell5.StyleIndex = 4U;
                cell5.CellValue = new CellValue("");
                string rowIndex6 = row.RowIndex;
                Cell cell6;
                Cell previousCell6;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex6)))
                {
                    cell6 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex6));
                }
                else
                {
                    previousCell6 = null;

                    for (int i6 = 3; i6 > 0; i6--)
                    {
                        previousCell6 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i6) + rowIndex6));

                        if (previousCell6 != null)
                            break;
                    }

                    cell6 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(3) + rowIndex6 };

                    row.InsertAfter(cell6, previousCell6);
                }

                cell6.DataType = CellValues.String;
                cell6.StyleIndex = 4U;
                cell6.CellValue = new CellValue("");

                for (int i = 4; i <= this.reportColumns.Count; i++)
                {
                    string rowIndex2 = row.RowIndex;
                    Cell cell2;
                    Cell previousCell2;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2)))
                    {
                        cell2 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2));
                    }
                    else
                    {
                        previousCell2 = null;

                        for (int i1 = i; i1 > 0; i1--)
                        {
                            previousCell2 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex2));

                            if (previousCell2 != null)
                                break;
                        }

                        cell2 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(i) + rowIndex2 };

                        row.InsertAfter(cell2, previousCell2);
                    }

                    cell2.DataType = CellValues.String;
                    cell2.StyleIndex = 4U;
                    cell2.CellValue = new CellValue(this.reportColumns[i - 1].title);
                }

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
                row = (Row)sheetData.LastChild;
                string rowIndex3 = row.RowIndex;
                Cell cell3;
                Cell previousCell3;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex3)))
                {
                    cell3 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex3));
                }
                else
                {
                    previousCell3 = null;

                    for (int i3 = 1; i3 > 0; i3--)
                    {
                        previousCell3 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i3) + rowIndex3));

                        if (previousCell3 != null)
                            break;
                    }

                    cell3 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(1) + rowIndex3 };

                    row.InsertAfter(cell3, previousCell3);
                }

                cell3.DataType = CellValues.String;
                cell3.StyleIndex = 4U;
                cell3.CellValue = new CellValue(ReportsSource.AvailabilityOfAtmNetwork);
                string rowIndex = row.RowIndex;
                Cell cell;
                Cell previousCell;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex)))
                {
                    cell = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex));
                }
                else
                {
                    previousCell = null;

                    for (int i1 = 2; i1 > 0; i1--)
                    {
                        previousCell = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex));

                        if (previousCell != null)
                            break;
                    }

                    cell = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(2) + rowIndex };

                    row.InsertAfter(cell, previousCell);
                }

                cell.DataType = CellValues.String;
                cell.StyleIndex = 4U;
                cell.CellValue = new CellValue("");
                string rowIndex1 = row.RowIndex;
                Cell cell1;
                Cell previousCell1;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex1)))
                {
                    cell1 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex1));
                }
                else
                {
                    previousCell1 = null;

                    for (int i2 = 3; i2 > 0; i2--)
                    {
                        previousCell1 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i2) + rowIndex1));

                        if (previousCell1 != null)
                            break;
                    }

                    cell1 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(3) + rowIndex1 };

                    row.InsertAfter(cell1, previousCell1);
                }

                cell1.DataType = CellValues.String;
                cell1.StyleIndex = 4U;
                cell1.CellValue = new CellValue("");
                ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A3", "C3");
                uint AvrgRowIndex = row.RowIndex;
                AvrgRow = row;

                double TechAvailAverage = 0;
                double DispenseAvailAverage = 0;
                double CashInAvailAverage = 0;
                double PaymentAvailAverage = 0;
                double StatAccountAvailAverage = 0;
                double EncashAvailAverage = 0;
                double FuncAvailAverage = 0;
                double LockedFactorSum = 0;
                double LockedFactorSumForCashAccept = 0;
                double LockedFactorSumForCashDispense = 0;

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
                row = (Row)sheetData.LastChild;

                for (int i = 1; i <= this.reportColumns.Count; i++)
                {
                    string rowIndex2 = row.RowIndex;
                    Cell cell2;
                    Cell previousCell2;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2)))
                    {
                        cell2 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2));
                    }
                    else
                    {
                        previousCell2 = null;

                        for (int i1 = i; i1 > 0; i1--)
                        {
                            previousCell2 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex2));

                            if (previousCell2 != null)
                                break;
                        }

                        cell2 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(i) + rowIndex2 };

                        row.InsertAfter(cell2, previousCell2);
                    }

                    cell2.DataType = CellValues.String;
                    cell2.StyleIndex = 4U;
                    cell2.CellValue = new CellValue(this.reportColumns[i - 1].title);
                }

                for (int i = 0; i < Data.Count; i++)
                {
                    TechAvailAverage += Data[i].avail.TechAvail * Data[i].avail.LockedFactor;
                    if ((Data[i].atmInfo.CashDispense ?? "") == "1")
                    {
                        DispenseAvailAverage += Data[i].avail.DispenseAvail * Data[i].avail.LockedFactor;
                        LockedFactorSumForCashDispense += Data[i].avail.LockedFactor;
                    }
                    //else
                    //    DispenseAvailAverage += 1 * Data[i].avail.LockedFactor;
                    if ((Data[i].atmInfo.CashAccept ?? "") == "1")
                    {
                        CashInAvailAverage += Data[i].avail.CashInAvail * Data[i].avail.LockedFactor;
                        LockedFactorSumForCashAccept += Data[i].avail.LockedFactor;
                    }
                    //else
                    //    CashInAvailAverage += 1 * Data[i].avail.LockedFactor;
                    PaymentAvailAverage += Data[i].avail.PaymentAvail * Data[i].avail.LockedFactor;
                    StatAccountAvailAverage += Data[i].avail.StatAccountAvail * Data[i].avail.LockedFactor;
                    EncashAvailAverage += Data[i].avail.EncashAvail * Data[i].avail.LockedFactor;
                    FuncAvailAverage += Data[i].avail.FunctionalAvail * Data[i].avail.LockedFactor;
                    LockedFactorSum += Data[i].avail.LockedFactor;

                    sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                    row = (Row)sheetData.LastChild;
                    string rowIndex2 = row.RowIndex;
                    Cell cell2;
                    Cell previousCell2;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex2)))
                    {
                        cell2 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex2));
                    }
                    else
                    {
                        previousCell2 = null;

                        for (int i1 = 1; i1 > 0; i1--)
                        {
                            previousCell2 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex2));

                            if (previousCell2 != null)
                                break;
                        }

                        cell2 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(1) + rowIndex2 };

                        row.InsertAfter(cell2, previousCell2);
                    }

                    cell2.DataType = CellValues.String;
                    cell2.StyleIndex = 1U;
                    cell2.CellValue = new CellValue(Data[i].atmInfo.Vizname);
                    if (this.BankName=="RNCB")
                    {
                        string rowIndex7 = row.RowIndex;
                        Cell cell13;
                        Cell previousCell13;

                        if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex7)))
                        {
                            cell13 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex7));
                        }
                        else
                        {
                            previousCell13 = null;

                            for (int i8 = 2; i8 > 0; i8--)
                            {
                                previousCell13 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i8) + rowIndex7));

                                if (previousCell13 != null)
                                    break;
                            }

                            cell13 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(2) + rowIndex7 };

                            row.InsertAfter(cell13, previousCell13);
                        }

                        cell13.DataType = CellValues.String;
                        cell13.StyleIndex = 5U;
                        cell13.CellValue = new CellValue(this.GetGroupForAtm(this.data[i].atmInfo.Id));
                    }
                    else
                    {
                        string rowIndex7 = row.RowIndex;
                        Cell cell7;
                        Cell previousCell7;

                        if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex7)))
                        {
                            cell7 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex7));
                        }
                        else
                        {
                            previousCell7 = null;

                            for (int i2 = 2; i2 > 0; i2--)
                            {
                                previousCell7 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i2) + rowIndex7));

                                if (previousCell7 != null)
                                    break;
                            }

                            cell7 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(2) + rowIndex7 };

                            row.InsertAfter(cell7, previousCell7);
                        }

                        cell7.DataType = CellValues.String;
                        cell7.StyleIndex = 5U;
                        cell7.CellValue = new CellValue(Data[i].atmInfo.Region);
                    }
                    string rowIndex8 = row.RowIndex;
                    Cell cell8;
                    Cell previousCell8;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex8)))
                    {
                        cell8 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex8));
                    }
                    else
                    {
                        previousCell8 = null;

                        for (int i3 = 3; i3 > 0; i3--)
                        {
                            previousCell8 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i3) + rowIndex8));

                            if (previousCell8 != null)
                                break;
                        }

                        cell8 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(3) + rowIndex8 };

                        row.InsertAfter(cell8, previousCell8);
                    }

                    cell8.DataType = CellValues.String;
                    cell8.StyleIndex = 5U;
                    cell8.CellValue = new CellValue(this.data[i].atmInfo.GeoAddress);
                    string rowIndex9 = row.RowIndex;
                    Cell cell9;
                    Cell previousCell9;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(4) + rowIndex9)))
                    {
                        cell9 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(4) + rowIndex9));
                    }
                    else
                    {
                        previousCell9 = null;

                        for (int i4 = 4; i4 > 0; i4--)
                        {
                            previousCell9 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i4) + rowIndex9));

                            if (previousCell9 != null)
                                break;
                        }

                        cell9 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(4) + rowIndex9 };

                        row.InsertAfter(cell9, previousCell9);
                    }

                    cell9.DataType = CellValues.String;
                    cell9.StyleIndex = 5U;
                    cell9.CellValue = new CellValue(Data[i].avail.TechAvail.ToString("p0"));
                    if ((Data[i].atmInfo.CashDispense ?? "") == "1")
                    {
                        string rowIndex7 = row.RowIndex;
                        Cell cell10;
                        Cell previousCell10;

                        if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex7)))
                        {
                            cell10 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex7));
                        }
                        else
                        {
                            previousCell10 = null;

                            for (int i5 = 5; i5 > 0; i5--)
                            {
                                previousCell10 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i5) + rowIndex7));

                                if (previousCell10 != null)
                                    break;
                            }

                            cell10 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(5) + rowIndex7 };

                            row.InsertAfter(cell10, previousCell10);
                        }

                        cell10.DataType = CellValues.String;
                        cell10.StyleIndex = 5U;
                        cell10.CellValue = new CellValue(Data[i].avail.DispenseAvail.ToString("p0"));
                    }
                    else
                    {
                        string rowIndex7 = row.RowIndex;
                        Cell cell11;
                        Cell previousCell11;

                        if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex7)))
                        {
                            cell11 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex7));
                        }
                        else
                        {
                            previousCell11 = null;

                            for (int i6 = 5; i6 > 0; i6--)
                            {
                                previousCell11 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i6) + rowIndex7));

                                if (previousCell11 != null)
                                    break;
                            }

                            cell11 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(5) + rowIndex7 };

                            row.InsertAfter(cell11, previousCell11);
                        }

                        cell11.DataType = CellValues.String;
                        cell11.StyleIndex = 5U;
                        cell11.CellValue = new CellValue("-");
                    }
                    if ((Data[i].atmInfo.CashAccept ?? "") == "1")
                    {
                        string rowIndex7 = row.RowIndex;
                        Cell cell12;
                        Cell previousCell12;

                        if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex7)))
                        {
                            cell12 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex7));
                        }
                        else
                        {
                            previousCell12 = null;

                            for (int i7 = 6; i7 > 0; i7--)
                            {
                                previousCell12 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i7) + rowIndex7));

                                if (previousCell12 != null)
                                    break;
                            }

                            cell12 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(6) + rowIndex7 };

                            row.InsertAfter(cell12, previousCell12);
                        }

                        cell12.DataType = CellValues.String;
                        cell12.StyleIndex = 5U;
                        cell12.CellValue = new CellValue(Data[i].avail.CashInAvail.ToString("p0"));
                    }
                    else
                    {
                        string rowIndex7 = row.RowIndex;
                        Cell cell14;
                        Cell previousCell14;

                        if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex7)))
                        {
                            cell14 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex7));
                        }
                        else
                        {
                            previousCell14 = null;

                            for (int i9 = 6; i9 > 0; i9--)
                            {
                                previousCell14 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i9) + rowIndex7));

                                if (previousCell14 != null)
                                    break;
                            }

                            cell14 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(6) + rowIndex7 };

                            row.InsertAfter(cell14, previousCell14);
                        }

                        cell14.DataType = CellValues.String;
                        cell14.StyleIndex = 5U;
                        cell14.CellValue = new CellValue("-");
                    }
                    string rowIndex10 = row.RowIndex;
                    Cell cell15;
                    Cell previousCell15;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(7) + rowIndex10)))
                    {
                        cell15 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(7) + rowIndex10));
                    }
                    else
                    {
                        previousCell15 = null;

                        for (int i10 = 7; i10 > 0; i10--)
                        {
                            previousCell15 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i10) + rowIndex10));

                            if (previousCell15 != null)
                                break;
                        }

                        cell15 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(7) + rowIndex10 };

                        row.InsertAfter(cell15, previousCell15);
                    }

                    cell15.DataType = CellValues.String;
                    cell15.StyleIndex = 5U;
                    cell15.CellValue = new CellValue(Data[i].avail.PaymentAvail.ToString("p0"));
                    string rowIndex18 = row.RowIndex;
                    Cell cell23;
                    Cell previousCell23;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(8) + rowIndex18)))
                    {
                        cell23 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(8) + rowIndex18));
                    }
                    else
                    {
                        previousCell23 = null;

                        for (int i12 = 8; i12 > 0; i12--)
                        {
                            previousCell23 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i12) + rowIndex18));

                            if (previousCell23 != null)
                                break;
                        }

                        cell23 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(8) + rowIndex18 };

                        row.InsertAfter(cell23, previousCell23);
                    }

                    cell23.DataType = CellValues.String;
                    cell23.StyleIndex = 5U;
                    cell23.CellValue = new CellValue(Data[i].avail.StatAccountAvail.ToString("p0"));
                    string rowIndex11 = row.RowIndex;
                    Cell cell16;
                    Cell previousCell16;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(9) + rowIndex11)))
                    {
                        cell16 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(9) + rowIndex11));
                    }
                    else
                    {
                        previousCell16 = null;

                        for (int i11 = 9; i11 > 0; i11--)
                        {
                            previousCell16 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i11) + rowIndex11));

                            if (previousCell16 != null)
                                break;
                        }

                        cell16 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(9) + rowIndex11 };

                        row.InsertAfter(cell16, previousCell16);
                    }

                    cell16.DataType = CellValues.String;
                    cell16.StyleIndex = 5U;
                    cell16.CellValue = new CellValue(Data[i].avail.FunctionalAvail.ToString("p0"));
                    // M3Utils.ExcelHelper.CreateCell(row, 10, row.RowIndex, Data[i].avail.EncashAvail.ToString("p0"), CellValues.String, 5U);
                }

                string rowIndex12 = AvrgRowIndex.ToString();
                Cell cell17;
                Cell previousCell17;

                if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(4) + rowIndex12)))
                {
                    cell17 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(4) + rowIndex12));
                }
                else
                {
                    previousCell17 = null;

                    for (int i12 = 4; i12 > 0; i12--)
                    {
                        previousCell17 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i12) + rowIndex12));

                        if (previousCell17 != null)
                            break;
                    }

                    cell17 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(4) + rowIndex12 };

                    AvrgRow.InsertAfter(cell17, previousCell17);
                }

                cell17.DataType = CellValues.String;
                cell17.StyleIndex = 5U;
                cell17.CellValue = new CellValue((TechAvailAverage / LockedFactorSum).ToString("p2"));
                string rowIndex13 = AvrgRowIndex.ToString();
                Cell cell18;
                Cell previousCell18;

                if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex13)))
                {
                    cell18 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex13));
                }
                else
                {
                    previousCell18 = null;

                    for (int i13 = 5; i13 > 0; i13--)
                    {
                        previousCell18 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i13) + rowIndex13));

                        if (previousCell18 != null)
                            break;
                    }

                    cell18 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(5) + rowIndex13 };

                    AvrgRow.InsertAfter(cell18, previousCell18);
                }

                cell18.DataType = CellValues.String;
                cell18.StyleIndex = 5U;
                cell18.CellValue = new CellValue((DispenseAvailAverage / LockedFactorSumForCashDispense).ToString("p2"));
                string rowIndex14 = AvrgRowIndex.ToString();
                Cell cell19;
                Cell previousCell19;

                if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex14)))
                {
                    cell19 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex14));
                }
                else
                {
                    previousCell19 = null;

                    for (int i14 = 6; i14 > 0; i14--)
                    {
                        previousCell19 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i14) + rowIndex14));

                        if (previousCell19 != null)
                            break;
                    }

                    cell19 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(6) + rowIndex14 };

                    AvrgRow.InsertAfter(cell19, previousCell19);
                }

                cell19.DataType = CellValues.String;
                cell19.StyleIndex = 5U;
                cell19.CellValue = new CellValue((CashInAvailAverage / LockedFactorSumForCashAccept).ToString("p2"));
                string rowIndex15 = AvrgRowIndex.ToString();
                Cell cell20;
                Cell previousCell20;

                if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(7) + rowIndex15)))
                {
                    cell20 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(7) + rowIndex15));
                }
                else
                {
                    previousCell20 = null;

                    for (int i15 = 7; i15 > 0; i15--)
                    {
                        previousCell20 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i15) + rowIndex15));

                        if (previousCell20 != null)
                            break;
                    }

                    cell20 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(7) + rowIndex15 };

                    AvrgRow.InsertAfter(cell20, previousCell20);
                }

                cell20.DataType = CellValues.String;
                cell20.StyleIndex = 5U;
                cell20.CellValue = new CellValue((PaymentAvailAverage / LockedFactorSum).ToString("p2"));
                string rowIndex16 = AvrgRowIndex.ToString();
                Cell cell21;
                Cell previousCell21;

                if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(8) + rowIndex16)))
                {
                    cell21 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(8) + rowIndex16));
                }
                else
                {
                    previousCell21 = null;

                    for (int i16 = 8; i16 > 0; i16--)
                    {
                        previousCell21 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i16) + rowIndex16));

                        if (previousCell21 != null)
                            break;
                    }

                    cell21 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(8) + rowIndex16 };

                    AvrgRow.InsertAfter(cell21, previousCell21);
                }

                cell21.DataType = CellValues.String;
                cell21.StyleIndex = 5U;
                cell21.CellValue = new CellValue((StatAccountAvailAverage / LockedFactorSum).ToString("p2"));
                //   M3Utils.ExcelHelper.CreateCell(AvrgRow, 8, AvrgRowIndex.ToString(), (EncashAvailAverage / LockedFactorSum).ToString("p2"), CellValues.String, 5U);
                string rowIndex17 = AvrgRowIndex.ToString();
                Cell cell22;
                Cell previousCell22;

                if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(9) + rowIndex17)))
                {
                    cell22 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(9) + rowIndex17));
                }
                else
                {
                    previousCell22 = null;

                    for (int i17 = 9; i17 > 0; i17--)
                    {
                        previousCell22 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i17) + rowIndex17));

                        if (previousCell22 != null)
                            break;
                    }

                    cell22 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(9) + rowIndex17 };

                    AvrgRow.InsertAfter(cell22, previousCell22);
                }

                cell22.DataType = CellValues.String;
                cell22.StyleIndex = 5U;
                cell22.CellValue = new CellValue((FuncAvailAverage / LockedFactorSum).ToString("p2"));

                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                for (int i = 1; i <= this.reportColumns.Count; i++)
                {
                    string rowIndex2 = row.RowIndex;
                    Cell cell2;
                    Cell previousCell2;

                    if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2)))
                    {
                        cell2 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2));
                    }
                    else
                    {
                        previousCell2 = null;

                        for (int i1 = i; i1 > 0; i1--)
                        {
                            previousCell2 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex2));

                            if (previousCell2 != null)
                                break;
                        }

                        cell2 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(i) + rowIndex2 };

                        row.InsertAfter(cell2, previousCell2);
                    }

                    cell2.DataType = CellValues.String;
                    cell2.StyleIndex = 6U;
                    cell2.CellValue = new CellValue("");
                }
            }
            catch (Exception exp)
            {
                Log.Instance.Info(
                    this + ".CreateHeaderRow(...) exception:",
                    exp.Message,
                    exp.Source,
                    exp.StackTrace);
            }
        }

        private string GetGroupForAtm(string sAtmId)
        {
            try
            {
                for (int k = 0; k < this.AtmGroups.Count; k++)
                {
                    if (this.AtmGroups[k].atmIds.Contains(sAtmId))
                    {
                        return this.AtmGroups[k].name;
                    }
                }
            }
            catch (Exception e)
            {
                Log.Instance.Info("GetGroupForAtm ReportAvailabilities(...) exception:");
                Log.Instance.Info(e.Message);
                Log.Instance.Info(e.Source);
                Log.Instance.Info(e.StackTrace);
            }
            return "Неизвестна";
        }

        public void SearchAvails()
        {
            List<Notch> Notches = new List<Notch>();
            List<Notch> OldNotches = new List<Notch>();
            Notch start = new Notch() { time = this.FromDate, type = NotchType.end, deleted = false };
            Notch end = new Notch() { time = this.ToDate, type = NotchType.begin, deleted = false };
            AvailsTime avails;
            Info info;
            List<Incident> IncidentList;

            try
            {
                foreach (string AtmId in this.AtmsId)
                {
                    double LockFactor = 1;
                    this.InitStatisticPeriod();
                    Notches.Clear();
                    Notches.Add(start);
                    Notches.Add(end);
                    Notches[0].deleted = false;
                    Notches[1].deleted = false;

                    avails = new AvailsTime();
                    avails.AtmId = AtmId;

                    try
                    {
                        info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                    }
                    catch
                    {
                        info = new Info();
                    }

                    var IncidentsForOneAtm = from incidents in this.Data.Incidents
                                             where incidents.atmId == AtmId
                                             orderby DateTime.Parse(incidents.timeCreated)
                                             select incidents;


                    #region IncidentsForOneAtm.IsNullOrEmpty
                    if (!IncidentsForOneAtm.IsNullOrEmpty())
                    {
                        #region LockedTime
                        var Locked = from incident in IncidentsForOneAtm
                                     where incident.typeId == this.LockedId
                                     select incident;

                        if (!Locked.IsNullOrEmpty())
                        {

                            foreach (Incident incident in Locked)
                            {
                                this.SearchNotches(Notches, incident);
                            }

                            if (!Notches.Any())
                            // банкомат не работал в указанный период
                            {
                                info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                                this.data.Add(new ReportAvailData()
                                {
                                    avail = new Avail()
                                    {
                                        DispenseAvail = 0,
                                        CashInAvail = 0,
                                        TechAvail = 0,
                                        EncashAvail = 0,
                                        PaymentAvail = 0,
                                        StatAccountAvail = 0,
                                        FunctionalAvail = 0,
                                        LockedFactor = LockFactor

                                    },
                                    atmInfo = info
                                });
                                continue;
                            }
                            avails.LockInterval = Notches.GetAvail();
                            LockFactor = avails.LockInterval.TotalSeconds / this.StatisticsPeriod.TotalSeconds;
                            this.StatisticsPeriod = avails.LockInterval;
                        }
                        #endregion

                        var AtmWorkIncidents = from incidents in IncidentsForOneAtm
                                               where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["AgentComm"]
                                               orderby DateTime.Parse(incidents.timeCreated)
                                               select incidents;


                        if (!AtmWorkIncidents.IsNullOrEmpty())
                        {
                            IncidentList = AtmWorkIncidents.ToList();
                            for (int i = 0; i < IncidentList.Count; i++) this.SearchNotches(Notches, IncidentList[i]);
                            if (!Notches.Any())
                            // банкомат не работал в указанный период
                            {
                                info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                                this.data.Add(new ReportAvailData()
                                {
                                    avail = new Avail()
                                    {
                                        DispenseAvail = 0,
                                        CashInAvail = 0,
                                        TechAvail = 0,
                                        EncashAvail = 0,
                                        PaymentAvail = 0,
                                        StatAccountAvail = 0,
                                        FunctionalAvail = 0,
                                        LockedFactor = LockFactor
                                    },
                                    atmInfo = info
                                });
                                continue;
                            }
                        }

                        var AtmConnectedIncidents = from incidents in IncidentsForOneAtm
                                                    where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["Communications"]
                                                    orderby DateTime.Parse(incidents.timeCreated)
                                                    select incidents;


                        if (!AtmConnectedIncidents.IsNullOrEmpty())
                        {
                            foreach (Incident incident in AtmConnectedIncidents) this.SearchNotches(Notches, incident);
                            if (!Notches.Any())
                            // банкомат не работал в указанный период
                            {
                                info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                                this.data.Add(new ReportAvailData()
                                {
                                    avail = new Avail()
                                    {
                                        DispenseAvail = 0,
                                        CashInAvail = 0,
                                        TechAvail = 0,
                                        EncashAvail = 0,
                                        PaymentAvail = 0,
                                        StatAccountAvail = 0,
                                        FunctionalAvail = 0,
                                        LockedFactor = LockFactor
                                    },
                                    atmInfo = info
                                });
                                continue;
                            }
                        }


                        //Считаем техническую доступность
                        avails.TechAvail = Notches.GetAvail();

                        // сохраняем состояние засечек, потом его восстановим
                        OldNotches.Clear();
                        OldNotches.AddRange(Notches);

                        var AtmCassetteEmptyFailure = from incidents in IncidentsForOneAtm
                                                      where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["Cassette_Status"]
                                                      orderby DateTime.Parse(incidents.timeCreated)
                                                      select incidents;

                        if (!AtmCassetteEmptyFailure.IsNullOrEmpty())
                        {
                            foreach (Incident incident in AtmCassetteEmptyFailure) this.SearchNotches(Notches, incident);
                        }

                        // посчитали инкассационную доступность  
                        avails.EncashAvail = Notches.GetAvail();

                        Notches.Clear();
                        Notches.AddRange(OldNotches); // восстановили состояние засечек

                        var AtmCardReaderFailure = from incidents in IncidentsForOneAtm
                                                   where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["CardReader"]
                                                   orderby DateTime.Parse(incidents.timeCreated)
                                                   select incidents;

                        if (!AtmCardReaderFailure.IsNullOrEmpty())
                        {
                            foreach (Incident incident in AtmCardReaderFailure) this.SearchNotches(Notches, incident);

                            if (!Notches.Any())
                            {
                                info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                                this.data.Add(new ReportAvailData()
                                {
                                    avail = new Avail()
                                    {
                                        DispenseAvail = 0,
                                        CashInAvail = 0,
                                        TechAvail = avails.TechAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                                        EncashAvail = avails.EncashAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                                        PaymentAvail = 0,
                                        StatAccountAvail = 0,
                                        FunctionalAvail = 0,
                                        LockedFactor = LockFactor
                                    },
                                    atmInfo = info
                                });
                                continue;
                            }
                        }

                        var AtmEncriptorFailure = from incidents in IncidentsForOneAtm
                                                  where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["EPP"]
                                                  orderby DateTime.Parse(incidents.timeCreated)
                                                  select incidents;

                        if (!AtmEncriptorFailure.IsNullOrEmpty())
                        {
                            foreach (Incident incident in AtmEncriptorFailure) this.SearchNotches(Notches, incident);

                            if (!Notches.Any())
                            {
                                info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                                this.data.Add(new ReportAvailData()
                                {
                                    avail = new Avail()
                                    {
                                        DispenseAvail = 0,
                                        CashInAvail = 0,
                                        TechAvail = avails.TechAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                                        EncashAvail = avails.EncashAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                                        PaymentAvail = 0,
                                        StatAccountAvail = 0,
                                        FunctionalAvail = 0,
                                        LockedFactor = LockFactor
                                    },
                                    atmInfo = info
                                });
                                continue;
                            }
                        }

                        IOrderedEnumerable<Incident> AtmJournalPrinterFailure;

                        if (this.BankName == "BKS")
                            AtmJournalPrinterFailure = null;
                        else
                        {
                            AtmJournalPrinterFailure = from incidents in IncidentsForOneAtm
                                                       where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["JournalPrinter"]
                                                       orderby DateTime.Parse(incidents.timeCreated)
                                                       select incidents;
                        }

                        if (!AtmJournalPrinterFailure.IsNullOrEmpty())
                        {
                            foreach (Incident incident in AtmJournalPrinterFailure) this.SearchNotches(Notches, incident);

                            if (!Notches.Any())
                            {
                                info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                                this.data.Add(new ReportAvailData()
                                {
                                    avail = new Avail()
                                    {
                                        DispenseAvail = 0,
                                        CashInAvail = 0,
                                        TechAvail = avails.TechAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                                        EncashAvail = avails.EncashAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                                        PaymentAvail = 0,
                                        StatAccountAvail = 0,
                                        FunctionalAvail = 0,
                                        LockedFactor = LockFactor
                                    },
                                    atmInfo = info
                                });
                                continue;
                            }
                        }

                        //доступность проведения платежей
                        avails.PaymentAvail = Notches.GetAvail();

                        OldNotches.Clear();
                        OldNotches.AddRange(Notches); // сохраняем состояние засечек.Восстановим для доступности выдачи наличных

                        var AtmCheckPrinterFailure = from incidents in IncidentsForOneAtm
                                                     where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["ReceiptPrinter"]
                                                     orderby DateTime.Parse(incidents.timeCreated)
                                                     select incidents;

                        if (!AtmCheckPrinterFailure.IsNullOrEmpty())
                        {
                            foreach (Incident incident in AtmCheckPrinterFailure) this.SearchNotches(Notches, incident);
                        }

                        if (M3UserSession.BankName == "RNCB")
                        {
                            //доступность проведения платежей
                            avails.PaymentAvail = Notches.GetAvail();
                        }

                        //доступность получения выписок
                        avails.StatAccountAvail = Notches.GetAvail();

                        List<Incident> AtmBNAFailure = null;
                        if ((info.CashAccept ?? "") == "1")
                        {
                            AtmBNAFailure = (from incidents in IncidentsForOneAtm
                                             where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["BNA"]
                                             orderby DateTime.Parse(incidents.timeCreated)
                                             select incidents).ToList();

                            if (!AtmBNAFailure.IsNullOrEmpty())
                            {
                                foreach (Incident incident in AtmBNAFailure) this.SearchNotches(Notches, incident);
                            }

                            //доступность приёма наличных
                            avails.CashInAvail = Notches.GetAvail();
                        }
                        else
                        {
                            avails.CashInAvail = new TimeSpan(0, 0, 0, 0, 0);
                        }



                        Notches.Clear();
                        Notches.AddRange(OldNotches);

                        if ((info.CashDispense ?? "") == "1")
                        {
                            var AtmDispenserFailure = from incidents in IncidentsForOneAtm
                                                      where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["Dispenser"]
                                                      orderby DateTime.Parse(incidents.timeCreated)
                                                      select incidents;

                            if (!AtmDispenserFailure.IsNullOrEmpty())
                            {
                                foreach (Incident incident in AtmDispenserFailure) this.SearchNotches(Notches, incident);

                                if (Notches.Any() && !AtmCassetteEmptyFailure.IsNullOrEmpty())
                                {
                                    foreach (Incident incident in AtmCassetteEmptyFailure) this.SearchNotches(Notches, incident);
                                }
                            }

                            //Доступность выдачи наличных
                            avails.DispenseAvail = Notches.GetAvail();

                            if (!AtmBNAFailure.IsNullOrEmpty())
                            {
                                foreach (Incident incident in AtmBNAFailure) this.SearchNotches(Notches, incident);
                                if (Notches.Any() && !AtmCheckPrinterFailure.IsNullOrEmpty())
                                {
                                    foreach (Incident incident in AtmCheckPrinterFailure) this.SearchNotches(Notches, incident);
                                }
                            }

                            //Combi availability
                            avails.CombiAvail = Notches.GetAvail();

                        }
                        else
                        {
                            avails.DispenseAvail = new TimeSpan(0, 0, 0, 0, 0);
                            avails.CombiAvail = avails.DispenseAvail;
                        }



                    }
                    else
                    {
                        // инцидентов по банкомату за указанный период нет в базе
                        avails.CashInAvail = this.StatisticsPeriod;
                        avails.DispenseAvail = this.StatisticsPeriod;
                        avails.EncashAvail = this.StatisticsPeriod;
                        avails.PaymentAvail = this.StatisticsPeriod;
                        avails.StatAccountAvail = this.StatisticsPeriod;
                        avails.TechAvail = this.StatisticsPeriod;
                        avails.CombiAvail = this.StatisticsPeriod;
                    }
                    #endregion

                    this.data.Add(new ReportAvailData()
                    {
                        avail = new Avail()
                        {
                            //AtmId = avails.AtmId,
                            DispenseAvail = avails.DispenseAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                            CashInAvail = avails.CashInAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                            TechAvail = avails.TechAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                            EncashAvail = avails.EncashAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                            PaymentAvail = avails.PaymentAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                            StatAccountAvail = avails.StatAccountAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                            FunctionalAvail = ((double)this.hasBna / this.AtmsId.Count) * (avails.CashInAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds) +
                                              ((double)this.hasDispenser / this.AtmsId.Count) * (avails.DispenseAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds) +
                                              ((double)this.hasBNAandDispenser / this.AtmsId.Count) * (avails.CombiAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds),
                            LockedFactor = LockFactor


                        },
                        atmInfo = info
                    });

                }
            }
            catch (Exception exp)
            {
                Log.Instance.Info(this + ".SearchAvails() exeption: " + exp.Message);
            }
        }

        public void SearchAvailsForAbsolute()
        {
            List<Notch> Notches = new List<Notch>();
            List<Notch> OldNotches = new List<Notch>();
            Notch start = new Notch() { time = this.FromDate, type = NotchType.end, deleted = false };
            Notch end = new Notch() { time = this.ToDate, type = NotchType.begin, deleted = false };
            AvailsTime avails;
            Info info;
            double LockFactor = 1;
            List<Incident> IncidentList;

            foreach (string AtmId in this.AtmsId)
            {
                Notches.Clear();
                Notches.Add(start);
                Notches.Add(end);
                Notches[0].deleted = false;
                Notches[1].deleted = false;

                avails = new AvailsTime();
                avails.AtmId = AtmId;

                try
                {
                    info = this.Data.AtmInfo.Where(f => f.Id == AtmId).First();
                }
                catch
                {
                    info = new Info();
                }

                var IncidentsForOneAtm = from incidents in this.Data.Incidents
                                         where incidents.atmId == AtmId
                                         orderby DateTime.Parse(incidents.timeCreated)
                                         select incidents;

                if (!IncidentsForOneAtm.IsNullOrEmpty())
                {

                    var AtmWorkIncidents = from incidents in IncidentsForOneAtm
                                           where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["AgentComm"]
                                           orderby DateTime.Parse(incidents.timeCreated)
                                           select incidents;

                    if (!AtmWorkIncidents.IsNullOrEmpty())
                    {
                        IncidentList = AtmWorkIncidents.ToList();

                        for (int i = 0; i < IncidentList.Count; i++) this.SearchNotches(Notches, IncidentList[i]);
                        if (!Notches.Any())
                        // банкомат не работал в указанный период
                        {
                            info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                            this.data.Add(new ReportAvailData()
                            {
                                avail = new Avail()
                                {
                                    DispenseAvail = 0,
                                    CashInAvail = 0,
                                    TechAvail = 0,
                                    EncashAvail = 0,
                                    PaymentAvail = 0,
                                    StatAccountAvail = 0,
                                    FunctionalAvail = 0,
                                    LockedFactor = LockFactor
                                },
                                atmInfo = info
                            });
                            continue;
                        }
                    }

                    avails.TechAvail = Notches.GetAvail();
                    //for client use avail
                    var InOutServiceIncidents = from incidents in IncidentsForOneAtm
                                                where incidents.deviceTypeId == "InOutService"
                                                orderby DateTime.Parse(incidents.timeCreated)
                                                select incidents;
                    if (!InOutServiceIncidents.IsNullOrEmpty())
                    {
                        foreach (Incident inc in InOutServiceIncidents)
                        {
                            this.SearchNotches(Notches, inc);
                        }

                        if (!Notches.Any())
                        // банкомат не работал в указанный период
                        {
                            info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                            this.data.Add(new ReportAvailData()
                            {
                                avail = new Avail()
                                {
                                    DispenseAvail = 0,
                                    CashInAvail = 0,
                                    TechAvail = avails.TechAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                                    EncashAvail = 0,
                                    PaymentAvail = 0,
                                    StatAccountAvail = 0,
                                    FunctionalAvail = 0,
                                    LockedFactor = LockFactor
                                },
                                atmInfo = info
                            });
                            continue;
                        }
                    }

                    avails.LockInterval = Notches.GetAvail();

                    var AtmJournalPrinterFailure = from incidents in IncidentsForOneAtm
                                                   where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["JournalPrinter"]
                                                   orderby DateTime.Parse(incidents.timeCreated)
                                                   select incidents;
                    if (!AtmJournalPrinterFailure.IsNullOrEmpty())
                    {
                        foreach (Incident incident in AtmJournalPrinterFailure) this.SearchNotches(Notches, incident);

                        if (!Notches.Any())
                        {
                            info = this.Data.AtmInfo.First(f => f.Id == AtmId);
                            this.data.Add(new ReportAvailData()
                            {
                                avail = new Avail()
                                {
                                    DispenseAvail = 0,
                                    CashInAvail = 0,
                                    TechAvail = avails.TechAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                                    EncashAvail = 0,
                                    PaymentAvail = 0,
                                    StatAccountAvail = 0,
                                    FunctionalAvail = avails.LockInterval.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                                    LockedFactor = LockFactor
                                },
                                atmInfo = info
                            });
                            continue;
                        }
                    }

                    OldNotches.Clear();
                    OldNotches.AddRange(Notches); // сохраняем состояние засечек.

                    if ((info.CashDispense ?? "") == "1")
                    {
                        var AtmDispenserFailure = from incidents in IncidentsForOneAtm
                                                  where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["Dispenser"]
                                                  orderby DateTime.Parse(incidents.timeCreated)
                                                  select incidents;
                        if (!AtmDispenserFailure.IsNullOrEmpty())
                        {
                            foreach (Incident incident in AtmDispenserFailure) this.SearchNotches(Notches, incident);
                        }
                        //Доступность выдачи наличных
                        avails.DispenseAvail = Notches.GetAvail();
                    }
                    else
                    {
                        avails.DispenseAvail = new TimeSpan(0, 0, 0, 0, 0);
                    }

                    Notches.Clear();
                    Notches.AddRange(OldNotches); // восстановили состояние засечек

                    var AtmCheckPrinterFailure = from incidents in IncidentsForOneAtm
                                                 where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["ReceiptPrinter"]
                                                 orderby DateTime.Parse(incidents.timeCreated)
                                                 select incidents;
                    if (!AtmCheckPrinterFailure.IsNullOrEmpty())
                    {
                        foreach (Incident incident in AtmCheckPrinterFailure) this.SearchNotches(Notches, incident);
                    }

                    avails.StatAccountAvail = Notches.GetAvail();
                    avails.PaymentAvail = avails.StatAccountAvail;

                    List<Incident> AtmBNAFailure;
                    if ((info.CashAccept ?? "") == "1")
                    {
                        AtmBNAFailure = (from incidents in IncidentsForOneAtm
                                         where Convert.ToInt32(incidents.deviceTypeId) == this.failuresDict["BNA"]
                                         orderby DateTime.Parse(incidents.timeCreated)
                                         select incidents).ToList();
                        if (!AtmBNAFailure.IsNullOrEmpty())
                        {
                            foreach (Incident incident in AtmBNAFailure) this.SearchNotches(Notches, incident);
                        }
                        //доступность приёма наличных
                        avails.CashInAvail = Notches.GetAvail();
                    }
                    else
                    {
                        avails.CashInAvail = new TimeSpan(0, 0, 0, 0, 0);
                    }
                }
                else
                {
                    // инцидентов по банкомату за указанный период нет в базе
                    avails.CashInAvail = this.StatisticsPeriod;
                    avails.DispenseAvail = this.StatisticsPeriod;
                    avails.EncashAvail = this.StatisticsPeriod;
                    avails.PaymentAvail = this.StatisticsPeriod;
                    avails.StatAccountAvail = this.StatisticsPeriod;
                    avails.TechAvail = this.StatisticsPeriod;
                    avails.CombiAvail = this.StatisticsPeriod;
                    avails.LockInterval = this.StatisticsPeriod;
                }

                this.data.Add(new ReportAvailData()
                {
                    avail = new Avail()
                    {
                        //AtmId = avails.AtmId,
                        DispenseAvail = avails.DispenseAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                        CashInAvail = avails.CashInAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                        TechAvail = avails.TechAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                        EncashAvail = avails.EncashAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                        PaymentAvail = avails.PaymentAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                        StatAccountAvail = avails.StatAccountAvail.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                        FunctionalAvail = avails.LockInterval.TotalSeconds / this.StatisticsPeriod.TotalSeconds,
                        LockedFactor = LockFactor
                    },
                    atmInfo = info
                });
            }// foreach atm
        }

        private void CreateDataRowsForAbsolute(WorksheetPart worksheetPart, List<ReportAvailData> Data)
        {
            Row row;
            Row AvrgRow;
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.First();

            row = (Row)sheetData.LastChild;

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            string rowIndex = row.RowIndex;
            Cell cell;
            Cell previousCell;

            if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex)))
            {
                cell = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex));
            }
            else
            {
                previousCell = null;

                for (int i1 = 1; i1 > 0; i1--)
                {
                    previousCell = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex));

                    if (previousCell != null)
                        break;
                }

                cell = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(1) + rowIndex };

                row.InsertAfter(cell, previousCell);
            }

            cell.DataType = CellValues.String;
            cell.StyleIndex = 4U;
            cell.CellValue = new CellValue("");
            string rowIndex1 = row.RowIndex;
            Cell cell1;
            Cell previousCell1;

            if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex1)))
            {
                cell1 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex1));
            }
            else
            {
                previousCell1 = null;

                for (int i2 = 2; i2 > 0; i2--)
                {
                    previousCell1 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i2) + rowIndex1));

                    if (previousCell1 != null)
                        break;
                }

                cell1 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(2) + rowIndex1 };

                row.InsertAfter(cell1, previousCell1);
            }

            cell1.DataType = CellValues.String;
            cell1.StyleIndex = 4U;
            cell1.CellValue = new CellValue("");
            for (int i = 3; i <= this.reportColumns.Count; i++)
            {
                string rowIndex2 = row.RowIndex;
                Cell cell2;
                Cell previousCell2;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2)))
                {
                    cell2 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2));
                }
                else
                {
                    previousCell2 = null;

                    for (int i1 = i; i1 > 0; i1--)
                    {
                        previousCell2 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex2));

                        if (previousCell2 != null)
                            break;
                    }

                    cell2 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(i) + rowIndex2 };

                    row.InsertAfter(cell2, previousCell2);
                }

                cell2.DataType = CellValues.String;
                cell2.StyleIndex = 4U;
                cell2.CellValue = new CellValue(this.reportColumns[i - 1].title);
            }

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;
            string rowIndex3 = row.RowIndex;
            Cell cell3;
            Cell previousCell3;

            if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex3)))
            {
                cell3 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex3));
            }
            else
            {
                previousCell3 = null;

                for (int i3 = 1; i3 > 0; i3--)
                {
                    previousCell3 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i3) + rowIndex3));

                    if (previousCell3 != null)
                        break;
                }

                cell3 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(1) + rowIndex3 };

                row.InsertAfter(cell3, previousCell3);
            }

            cell3.DataType = CellValues.String;
            cell3.StyleIndex = 4U;
            cell3.CellValue = new CellValue("Доступность сети банкоматов");
            string rowIndex4 = row.RowIndex;
            Cell cell4;
            Cell previousCell4;

            if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex4)))
            {
                cell4 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex4));
            }
            else
            {
                previousCell4 = null;

                for (int i4 = 2; i4 > 0; i4--)
                {
                    previousCell4 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i4) + rowIndex4));

                    if (previousCell4 != null)
                        break;
                }

                cell4 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(2) + rowIndex4 };

                row.InsertAfter(cell4, previousCell4);
            }

            cell4.DataType = CellValues.String;
            cell4.StyleIndex = 4U;
            cell4.CellValue = new CellValue("");
            string rowIndex5 = row.RowIndex;
            Cell cell5;
            Cell previousCell5;

            if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex5)))
            {
                cell5 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex5));
            }
            else
            {
                previousCell5 = null;

                for (int i5 = 3; i5 > 0; i5--)
                {
                    previousCell5 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i5) + rowIndex5));

                    if (previousCell5 != null)
                        break;
                }

                cell5 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(3) + rowIndex5 };

                row.InsertAfter(cell5, previousCell5);
            }

            cell5.DataType = CellValues.String;
            cell5.StyleIndex = 4U;
            cell5.CellValue = new CellValue("");
            ExcelHelper.MergeCellsInRange(worksheetPart.Worksheet, "A3", "C3");
            uint AvrgRowIndex = row.RowIndex;
            AvrgRow = row;

            double TechAvailAverage = 0;
            double DispenseAvailAverage = 0;
            double CashInAvailAverage = 0;
            double PaymentAvailAverage = 0;
            double StatAccountAvailAverage = 0;
            double EncashAvailAverage = 0;
            double FuncAvailAverage = 0;
            double LockedFactorSum = 0;
            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1), Height = 20D, CustomHeight = true });
            row = (Row)sheetData.LastChild;

            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                string rowIndex2 = row.RowIndex;
                Cell cell2;
                Cell previousCell2;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2)))
                {
                    cell2 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2));
                }
                else
                {
                    previousCell2 = null;

                    for (int i1 = i; i1 > 0; i1--)
                    {
                        previousCell2 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex2));

                        if (previousCell2 != null)
                            break;
                    }

                    cell2 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(i) + rowIndex2 };

                    row.InsertAfter(cell2, previousCell2);
                }

                cell2.DataType = CellValues.String;
                cell2.StyleIndex = 4U;
                cell2.CellValue = new CellValue(this.reportColumns[i - 1].title);
            }

            for (int i = 0; i < Data.Count; i++)
            {
                TechAvailAverage += Data[i].avail.TechAvail;
                DispenseAvailAverage += Data[i].avail.DispenseAvail;
                CashInAvailAverage += Data[i].avail.CashInAvail;
                PaymentAvailAverage += Data[i].avail.PaymentAvail;
                StatAccountAvailAverage += Data[i].avail.StatAccountAvail;
                EncashAvailAverage += Data[i].avail.EncashAvail;
                FuncAvailAverage += Data[i].avail.FunctionalAvail;
                LockedFactorSum += Data[i].avail.LockedFactor;


                sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
                row = (Row)sheetData.LastChild;
                string rowIndex2 = row.RowIndex;
                Cell cell2;
                Cell previousCell2;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex2)))
                {
                    cell2 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(1) + rowIndex2));
                }
                else
                {
                    previousCell2 = null;

                    for (int i1 = 1; i1 > 0; i1--)
                    {
                        previousCell2 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex2));

                        if (previousCell2 != null)
                            break;
                    }

                    cell2 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(1) + rowIndex2 };

                    row.InsertAfter(cell2, previousCell2);
                }

                cell2.DataType = CellValues.String;
                cell2.StyleIndex = 1U;
                cell2.CellValue = new CellValue(Data[i].atmInfo.Vizname);
                string rowIndex6 = row.RowIndex;
                Cell cell6;
                Cell previousCell6;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex6)))
                {
                    cell6 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(2) + rowIndex6));
                }
                else
                {
                    previousCell6 = null;

                    for (int i2 = 2; i2 > 0; i2--)
                    {
                        previousCell6 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i2) + rowIndex6));

                        if (previousCell6 != null)
                            break;
                    }

                    cell6 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(2) + rowIndex6 };

                    row.InsertAfter(cell6, previousCell6);
                }

                cell6.DataType = CellValues.String;
                cell6.StyleIndex = 5U;
                cell6.CellValue = new CellValue(this.data[i].atmInfo.Region);
                string rowIndex7 = row.RowIndex;
                Cell cell7;
                Cell previousCell7;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex7)))
                {
                    cell7 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(3) + rowIndex7));
                }
                else
                {
                    previousCell7 = null;

                    for (int i3 = 3; i3 > 0; i3--)
                    {
                        previousCell7 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i3) + rowIndex7));

                        if (previousCell7 != null)
                            break;
                    }

                    cell7 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(3) + rowIndex7 };

                    row.InsertAfter(cell7, previousCell7);
                }

                cell7.DataType = CellValues.String;
                cell7.StyleIndex = 5U;
                cell7.CellValue = new CellValue(this.data[i].atmInfo.GeoAddress);
                string rowIndex8 = row.RowIndex;
                Cell cell8;
                Cell previousCell8;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(4) + rowIndex8)))
                {
                    cell8 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(4) + rowIndex8));
                }
                else
                {
                    previousCell8 = null;

                    for (int i4 = 4; i4 > 0; i4--)
                    {
                        previousCell8 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i4) + rowIndex8));

                        if (previousCell8 != null)
                            break;
                    }

                    cell8 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(4) + rowIndex8 };

                    row.InsertAfter(cell8, previousCell8);
                }

                cell8.DataType = CellValues.String;
                cell8.StyleIndex = 5U;
                cell8.CellValue = new CellValue(Data[i].avail.TechAvail.ToString("p0"));
                string rowIndex9 = row.RowIndex;
                Cell cell9;
                Cell previousCell9;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex9)))
                {
                    cell9 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex9));
                }
                else
                {
                    previousCell9 = null;

                    for (int i5 = 5; i5 > 0; i5--)
                    {
                        previousCell9 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i5) + rowIndex9));

                        if (previousCell9 != null)
                            break;
                    }

                    cell9 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(5) + rowIndex9 };

                    row.InsertAfter(cell9, previousCell9);
                }

                cell9.DataType = CellValues.String;
                cell9.StyleIndex = 5U;
                cell9.CellValue = new CellValue(Data[i].avail.DispenseAvail.ToString("p0"));
                string rowIndex10 = row.RowIndex;
                Cell cell10;
                Cell previousCell10;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex10)))
                {
                    cell10 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex10));
                }
                else
                {
                    previousCell10 = null;

                    for (int i6 = 6; i6 > 0; i6--)
                    {
                        previousCell10 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i6) + rowIndex10));

                        if (previousCell10 != null)
                            break;
                    }

                    cell10 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(6) + rowIndex10 };

                    row.InsertAfter(cell10, previousCell10);
                }

                cell10.DataType = CellValues.String;
                cell10.StyleIndex = 5U;
                cell10.CellValue = new CellValue(Data[i].avail.CashInAvail.ToString("p0"));
                string rowIndex11 = row.RowIndex;
                Cell cell11;
                Cell previousCell11;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(7) + rowIndex11)))
                {
                    cell11 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(7) + rowIndex11));
                }
                else
                {
                    previousCell11 = null;

                    for (int i7 = 7; i7 > 0; i7--)
                    {
                        previousCell11 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i7) + rowIndex11));

                        if (previousCell11 != null)
                            break;
                    }

                    cell11 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(7) + rowIndex11 };

                    row.InsertAfter(cell11, previousCell11);
                }

                cell11.DataType = CellValues.String;
                cell11.StyleIndex = 5U;
                cell11.CellValue = new CellValue(Data[i].avail.PaymentAvail.ToString("p0"));
                string rowIndex12 = row.RowIndex;
                Cell cell12;
                Cell previousCell12;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(8) + rowIndex12)))
                {
                    cell12 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(8) + rowIndex12));
                }
                else
                {
                    previousCell12 = null;

                    for (int i8 = 8; i8 > 0; i8--)
                    {
                        previousCell12 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i8) + rowIndex12));

                        if (previousCell12 != null)
                            break;
                    }

                    cell12 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(8) + rowIndex12 };

                    row.InsertAfter(cell12, previousCell12);
                }

                cell12.DataType = CellValues.String;
                cell12.StyleIndex = 5U;
                cell12.CellValue = new CellValue(Data[i].avail.StatAccountAvail.ToString("p0"));
                string rowIndex13 = row.RowIndex;
                Cell cell13;
                Cell previousCell13;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(9) + rowIndex13)))
                {
                    cell13 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(9) + rowIndex13));
                }
                else
                {
                    previousCell13 = null;

                    for (int i9 = 9; i9 > 0; i9--)
                    {
                        previousCell13 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i9) + rowIndex13));

                        if (previousCell13 != null)
                            break;
                    }

                    cell13 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(9) + rowIndex13 };

                    row.InsertAfter(cell13, previousCell13);
                }

                cell13.DataType = CellValues.String;
                cell13.StyleIndex = 5U;
                cell13.CellValue = new CellValue(Data[i].avail.FunctionalAvail.ToString("p0"));
            }

            string rowIndex14 = AvrgRowIndex.ToString();
            Cell cell14;
            Cell previousCell14;

            if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(4) + rowIndex14)))
            {
                cell14 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(4) + rowIndex14));
            }
            else
            {
                previousCell14 = null;

                for (int i10 = 4; i10 > 0; i10--)
                {
                    previousCell14 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i10) + rowIndex14));

                    if (previousCell14 != null)
                        break;
                }

                cell14 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(4) + rowIndex14 };

                AvrgRow.InsertAfter(cell14, previousCell14);
            }

            cell14.DataType = CellValues.String;
            cell14.StyleIndex = 5U;
            cell14.CellValue = new CellValue((TechAvailAverage / LockedFactorSum).ToString("p0"));
            string rowIndex15 = AvrgRowIndex.ToString();
            Cell cell15;
            Cell previousCell15;

            if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex15)))
            {
                cell15 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(5) + rowIndex15));
            }
            else
            {
                previousCell15 = null;

                for (int i11 = 5; i11 > 0; i11--)
                {
                    previousCell15 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i11) + rowIndex15));

                    if (previousCell15 != null)
                        break;
                }

                cell15 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(5) + rowIndex15 };

                AvrgRow.InsertAfter(cell15, previousCell15);
            }

            cell15.DataType = CellValues.String;
            cell15.StyleIndex = 5U;
            cell15.CellValue = new CellValue((DispenseAvailAverage / LockedFactorSum).ToString("p0"));
            string rowIndex16 = AvrgRowIndex.ToString();
            Cell cell16;
            Cell previousCell16;

            if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex16)))
            {
                cell16 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(6) + rowIndex16));
            }
            else
            {
                previousCell16 = null;

                for (int i12 = 6; i12 > 0; i12--)
                {
                    previousCell16 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i12) + rowIndex16));

                    if (previousCell16 != null)
                        break;
                }

                cell16 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(6) + rowIndex16 };

                AvrgRow.InsertAfter(cell16, previousCell16);
            }

            cell16.DataType = CellValues.String;
            cell16.StyleIndex = 5U;
            cell16.CellValue = new CellValue((CashInAvailAverage / LockedFactorSum).ToString("p0"));
            string rowIndex17 = AvrgRowIndex.ToString();
            Cell cell17;
            Cell previousCell17;

            if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(7) + rowIndex17)))
            {
                cell17 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(7) + rowIndex17));
            }
            else
            {
                previousCell17 = null;

                for (int i13 = 7; i13 > 0; i13--)
                {
                    previousCell17 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i13) + rowIndex17));

                    if (previousCell17 != null)
                        break;
                }

                cell17 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(7) + rowIndex17 };

                AvrgRow.InsertAfter(cell17, previousCell17);
            }

            cell17.DataType = CellValues.String;
            cell17.StyleIndex = 5U;
            cell17.CellValue = new CellValue((PaymentAvailAverage / LockedFactorSum).ToString("p0"));
            string rowIndex18 = AvrgRowIndex.ToString();
            Cell cell18;
            Cell previousCell18;

            if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(8) + rowIndex18)))
            {
                cell18 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(8) + rowIndex18));
            }
            else
            {
                previousCell18 = null;

                for (int i14 = 8; i14 > 0; i14--)
                {
                    previousCell18 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i14) + rowIndex18));

                    if (previousCell18 != null)
                        break;
                }

                cell18 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(8) + rowIndex18 };

                AvrgRow.InsertAfter(cell18, previousCell18);
            }

            cell18.DataType = CellValues.String;
            cell18.StyleIndex = 5U;
            cell18.CellValue = new CellValue((StatAccountAvailAverage / LockedFactorSum).ToString("p0"));
            string rowIndex19 = AvrgRowIndex.ToString();
            Cell cell19;
            Cell previousCell19;

            if (AvrgRow.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(9) + rowIndex19)))
            {
                cell19 = AvrgRow.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(9) + rowIndex19));
            }
            else
            {
                previousCell19 = null;

                for (int i15 = 9; i15 > 0; i15--)
                {
                    previousCell19 = AvrgRow.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i15) + rowIndex19));

                    if (previousCell19 != null)
                        break;
                }

                cell19 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(9) + rowIndex19 };

                AvrgRow.InsertAfter(cell19, previousCell19);
            }

            cell19.DataType = CellValues.String;
            cell19.StyleIndex = 5U;
            cell19.CellValue = new CellValue((FuncAvailAverage / LockedFactorSum).ToString("p0"));

            sheetData.Append(new Row() { RowIndex = (row.RowIndex + 1) });
            row = (Row)sheetData.LastChild;
            for (int i = 1; i <= this.reportColumns.Count; i++)
            {
                string rowIndex2 = row.RowIndex;
                Cell cell2;
                Cell previousCell2;

                if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2)))
                {
                    cell2 = row.Elements<Cell>().First(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i) + rowIndex2));
                }
                else
                {
                    previousCell2 = null;

                    for (int i1 = i; i1 > 0; i1--)
                    {
                        previousCell2 = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ExcelHelper.ColumnNameByIndex(i1) + rowIndex2));

                        if (previousCell2 != null)
                            break;
                    }

                    cell2 = new Cell() { CellReference = ExcelHelper.ColumnNameByIndex(i) + rowIndex2 };

                    row.InsertAfter(cell2, previousCell2);
                }

                cell2.DataType = CellValues.String;
                cell2.StyleIndex = 6U;
                cell2.CellValue = new CellValue("");
            }
        }
    }
}