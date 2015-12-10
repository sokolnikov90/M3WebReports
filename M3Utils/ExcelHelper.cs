using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace M3Utils
{
    using System.Xml.Linq;

    public static class ExcelHelper
    {
        public static string ColumnNameByIndex(int columnIndex)
        {
            string columnName = "";

            while (columnIndex > 0)
            {
                var remainder = (columnIndex - 1) % 26;

                columnName = Convert.ToChar(65 + remainder).ToString() + columnName;

                columnIndex = (int)((columnIndex - remainder) / 26);
            }

            return columnName;
        }

        public static bool SetColumnWidth(Worksheet worksheet, int columnIndex, DoubleValue width)
        {
            Columns columns;

            Column column;
            Column previousColumn;

            columns = worksheet.Elements<Columns>().FirstOrDefault() ?? worksheet.InsertAt(new Columns(), 0);

            if (columns.Elements<Column>().All(item => item.Min != columnIndex))
            {
                previousColumn = null;

                for (int i = columnIndex; i > 0; i--)
                {
                    previousColumn = columns.Elements<Column>().FirstOrDefault(item => item.Min == i);

                    if (previousColumn != null)
                        break;
                }

                column = new Column
                             {
                                 Min = Convert.ToUInt32(columnIndex),
                                 Max = Convert.ToUInt32(columnIndex),
                                 Width = width,
                                 CustomWidth = true
                             };


                columns.InsertAfter(column, previousColumn);
            }

            return true;
        }

        public static void CreateCell(Row row, int columnIndex, string rowIndex, string value, CellValues valueType, UInt32Value styleIndex)
        {
            Cell cell;
            Cell previousCell;

            if (row.Elements<Cell>().Any(item => item.CellReference.Value == (ColumnNameByIndex(columnIndex) + rowIndex)))
            {
                cell = row.Elements<Cell>().First(item => item.CellReference.Value == (ColumnNameByIndex(columnIndex) + rowIndex));
            }
            else
            {
                previousCell = null;

                for (int i = columnIndex; i > 0; i--)
                {
                    previousCell = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == (ColumnNameByIndex(i) + rowIndex));

                    if (previousCell != null)
                        break;
                }

                cell = new Cell() { CellReference = ColumnNameByIndex(columnIndex) + rowIndex };

                row.InsertAfter(cell, previousCell);
            }

            cell.DataType = valueType;
            cell.StyleIndex = styleIndex;
            cell.CellValue = new CellValue(value);
        }

        public static void MergeCellsInRange(Worksheet worksheet, string cellFrom, string cellTo)
        {
            SheetData sheetData = (SheetData)worksheet.First();

            MergeCells mergeCells;

            if (worksheet.Elements<MergeCells>().Any())
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();

                worksheet.InsertAfter(mergeCells, sheetData);
            }

            MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cellFrom + ":" + cellTo) };

            mergeCells.Append(mergeCell);
        }

        public static Stylesheet MakeStyleSheet()
        {
            Stylesheet stylesheet = new Stylesheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" }
            };

            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            var fonts = MakeStyleSheet_Fonts();
            var fills = MakeStyleSheet_Fills();
            var borders = MakeStyleSheet_Borders();
            var cellFormats = MakeStyleSheet_CellFormats();

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellFormats);

            return stylesheet;
        }

        private static Fonts MakeStyleSheet_Fonts()
        {
            Fonts fonts = new Fonts() { Count = 2U, KnownFonts = true };

            // Font = 1U
            Font font11D = new Font();
            FontSize fontSize11D = new FontSize() { Val = 11D };
            font11D.Append(fontSize11D);


            // Font = 2U
            Font font11DBold = new Font();
            Bold bold11DBold = new Bold();
            FontSize fontSize11DBold = new FontSize() { Val = 11D };
            font11DBold.Append(bold11DBold);
            font11DBold.Append(fontSize11DBold);

            fonts.Append(font11D);
            fonts.Append(font11DBold);

            return fonts;
        }

        private static Fills MakeStyleSheet_Fills()
        {
            Fills fills = new Fills() { Count = (UInt32Value)5U };

            // FillId = 0
            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);

            // FillId = 1
            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
            fill2.Append(patternFill2);

            // FillId = 2,RED
            Fill fill3 = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFFF0000" };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);
            fill3.Append(patternFill3);

            // FillId = 3,Green
            Fill fill4 = new Fill();
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FF7CFC00" };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);
            fill4.Append(patternFill4);

            // FillId = 4,YELLOW
            Fill fill5 = new Fill();
            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFFFFF00" };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill5.Append(foregroundColor3);
            patternFill5.Append(backgroundColor3);
            fill5.Append(patternFill5);

            fills.Append(fill1);
            fills.Append(fill2);
            fills.Append(fill3);
            fills.Append(fill4);
            fills.Append(fill5);

            return fills;
        }

        private static Borders MakeStyleSheet_Borders()
        {
            Borders borders = new Borders() { Count = (UInt32Value)5U };

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

            return borders;
        }

        private static CellFormats MakeStyleSheet_CellFormats()
        {
            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)10U };

            // 0U - Default format.
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFont = true };

            // 1U - Default format with center-center alignment.
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, ApplyFont = true, ApplyAlignment = true };
            cellFormat2.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center });

            // 2U - Default format with center alignment.
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, ApplyFont = true, ApplyAlignment = true };
            cellFormat3.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true });

            // 3U - Bold font format.
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFont = true };

            // 4U - Bold font format with center alignment.
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, ApplyFont = true, ApplyAlignment = true };
            cellFormat5.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true });

            // 5U
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, ApplyFont = true, ApplyAlignment = true };
            cellFormat6.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true });

            // 6U
            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, ApplyFont = true, ApplyAlignment = true };
            cellFormat7.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center });

            // 7U Green background
            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)4U, ApplyFont = true, ApplyAlignment = true };
            cellFormat8.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true });

            // 9U Yellow background
            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)4U, ApplyFont = true, ApplyAlignment = true };
            cellFormat9.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true });

            // 8U Red background
            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)4U, ApplyFont = true, ApplyAlignment = true };
            cellFormat10.Append(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true });

            cellFormats.Append(cellFormat1);
            cellFormats.Append(cellFormat2);
            cellFormats.Append(cellFormat3);
            cellFormats.Append(cellFormat4);
            cellFormats.Append(cellFormat5);
            cellFormats.Append(cellFormat6);
            cellFormats.Append(cellFormat7);
            cellFormats.Append(cellFormat8);
            cellFormats.Append(cellFormat9);
            cellFormats.Append(cellFormat10);

            return cellFormats;
        }
    }
}