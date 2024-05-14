using System;
using DevExpress.Spreadsheet;
using DevExpress.Utils;
using System.Globalization;
using System.Collections.Generic;

namespace SpreadsheetControl_API 
{
    public static class SortActions {
        #region Actions
        public static Action<IWorkbook> SimpleSortAction = SimpleSortValue;
        public static Action<IWorkbook> DescendingOrderAction = DescendingOrderValue;
        public static Action<IWorkbook> SelectComparerAction = SelectComparerValue;
        public static Action<IWorkbook> SortBySpecifiedColumnAction = SortBySpecifiedColumnValue;
        public static Action<IWorkbook> SortByMultipleColumnsAction = SortByMultipleColumnsValue;
        public static Action<IWorkbook> SortByFillColorAction = SortByFillColorValue;
        public static Action<IWorkbook> SortByFontColorAction = SortByFontColorValue;

        #endregion

        static void SimpleSortValue(IWorkbook workbook) {
            #region #SimpleSort
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill in the range.
            worksheet.Cells["A2"].Value = "Donald Dozier Bradley";
            worksheet.Cells["A3"].Value = "Tony Charles Mccallum-Geteer";
            worksheet.Cells["A4"].Value = "Calvin Liu";
            worksheet.Cells["A5"].Value = "Anita A Boyd";
            worksheet.Cells["A6"].Value = "Angela R. Scott";
            worksheet.Cells["A7"].Value = "D Fox";

            // Sort the range in ascending order.
            CellRange range = worksheet.Range["A2:A7"];
            worksheet.Sort(range);

            // Create a heading.
            CellRange header = worksheet.Range["A1"];
            header[0].Value = "Ascending order";
            header.ColumnWidthInCharacters = 30;
            header.Style = workbook.Styles["Heading 1"];
            #endregion #SimpleSort
        }

        static void DescendingOrderValue(IWorkbook workbook) {
            #region #DescendingOrder
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill in the range.
            worksheet.Cells["A2"].Value = "Donald Dozier Bradley";
            worksheet.Cells["A3"].Value = "Tony Charles Mccallum-Geteer";
            worksheet.Cells["A4"].Value = "Calvin Liu";
            worksheet.Cells["A5"].Value = "Anita A Boyd";
            worksheet.Cells["A6"].Value = "Angela R. Scott";
            worksheet.Cells["A7"].Value = "D Fox";

            // Sort the range in descending order.
            CellRange range = worksheet.Range["A2:A7"];
            worksheet.Sort(range, false);

            // Create a heading.
            CellRange header = worksheet.Range["A1"];
            header[0].Value = "Descending order";
            header.ColumnWidthInCharacters = 30;
            header.Style = workbook.Styles["Heading 1"];
            #endregion #DescendingOrder
        }

        static void SelectComparerValue(IWorkbook workbook) {
            #region #SelectComparer
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill in the range.
            worksheet.Cells["A2"].Value = "Donald Dozier Bradley";
            worksheet.Cells["A3"].Value = "Tony Charles Mccallum-Geteer";
            worksheet.Cells["A4"].Value = "Calvin Liu";
            worksheet.Cells["A5"].Value = "Anita A Boyd";
            worksheet.Cells["A6"].Value = "Angela R. Scott";
            worksheet.Cells["A7"].Value = "D Fox";

            // Sort values using a custom comparer.
            CellRange range = worksheet.Range["A2:A7"];
            worksheet.Sort(range, 0, new SampleComparer());

            // Create a heading.
            CellRange header = worksheet.Range["A1"];
            header[0].Value = "Use a custom comparer";
            header.ColumnWidthInCharacters = 30;
            header.Style = workbook.Styles["Heading 1"];
            #endregion #SelectComparer
        }

        static void SortBySpecifiedColumnValue(IWorkbook workbook) {
            #region #SortBySpecifiedColumn
            workbook.LoadDocument("Documents\\Sortsample.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Sort by a column with offset = 6 in the range being sorted.
            // Use ascending order.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, 3);

            // Add a note.
            worksheet["D1"].Value = "Sort by column with index = 3 in ascending order";
            worksheet.Visible = true;
            #endregion #SortBySpecifiedColumn
        }

        static void SortByMultipleColumnsValue(IWorkbook workbook) {
            #region #SortByMultipleColumns
            workbook.LoadDocument("Documents\\Sortsample.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Create sorting fields.
            List<SortField> fields = new List<SortField>();
            
            // First sorting field. First column (offset = 0) will be sorted using ascending order.
            SortField sortField1 = new SortField();
            sortField1.ColumnOffset = 0;
            sortField1.Comparer = worksheet.Comparers.Ascending;
            fields.Add(sortField1);

            // Second sorting field. Second column (offset = 1) will be sorted using ascending order.
            SortField sortField2 = new SortField();
            sortField2.ColumnOffset = 1;
            sortField2.Comparer = worksheet.Comparers.Ascending;
            fields.Add(sortField2);
            
            // Sort the range by sorting fields.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, fields);

            #endregion #SortByMultipleColumns
            // Add a note.
            worksheet["D1"].Value = "Sort by two columns: first and second in ascending order";
        }

        static void SortByFillColorValue(IWorkbook workbook)
        {
            #region #SortByFillColor
            workbook.LoadDocument("Documents\\Sortsample.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Sort the "A3:F22" range by column "A" in ascending order.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, 0, worksheet["A3"].Fill);

            #endregion #SortByFillColor
        }

        static void SortByFontColorValue(IWorkbook workbook)
        {
            #region #SortByFontColor
            workbook.LoadDocument("Documents\\Sortsample.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Sort the "A3:F22" range by column "F" in ascending order.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, 5, worksheet["F12"].Font.Color);

            #endregion #SortByFontColor
        }

    }
}
