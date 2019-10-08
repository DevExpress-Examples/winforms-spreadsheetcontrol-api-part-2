using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetControl_API
{
    class GroupAndOutlineActions
    {
        #region Actions
        public static Action<IWorkbook> GroupRowsAction = GroupRowsValue;
        public static Action<IWorkbook> GroupColumnsAction = GroupColumnsValue;
        public static Action<IWorkbook> UngroupRowsAction = UngroupRowsValue;
        public static Action<IWorkbook> UngroupColumnsAction = UngroupColumnsValue;
        public static Action<IWorkbook> AutoOutlineAction = AutoOutlineValue;
        public static Action<IWorkbook> SubtotalAction = SubtotalValue;
        #endregion

            static void GroupRowsValue(IWorkbook workbook)
            {
                workbook.LoadDocument("Documents\\SalesReport.xlsx");
                workbook.BeginUpdate();
                try
                {
                    Worksheet worksheet = workbook.Worksheets["Sales Analysis"];
                    workbook.Worksheets.ActiveWorksheet = worksheet;

                    #region #GroupRows
                    // Group four rows starting from the third row and collapse the group.
                    worksheet.Rows.Group(2, 5, true);

                    // Group four rows starting from the ninth row and expand the group.
                    worksheet.Rows.Group(8, 11, false);

                    // Create the outer group of rows by grouping rows 2 through 13. 
                    worksheet.Rows.Group(1, 12, false);
                    #endregion #GroupRows
                }
                finally { workbook.EndUpdate(); }
            }

            static void GroupColumnsValue(IWorkbook workbook)
            {
                workbook.LoadDocument("Documents\\SalesReport.xlsx");
                workbook.BeginUpdate();
                try
                {
                    Worksheet worksheet = workbook.Worksheets["Sales Analysis"];
                    workbook.Worksheets.ActiveWorksheet = worksheet;

                    #region #GroupColumns
                    // Group four columns starting from the third column "C" and expand the group.
                    worksheet.Columns.Group(2, 5, false);
                    #endregion #GroupColumns
                }
                finally { workbook.EndUpdate(); }
            }

            static void UngroupRowsValue(IWorkbook workbook)
            {
                workbook.LoadDocument("Documents\\SalesReport.xlsx");
                workbook.BeginUpdate();
                try
                {
                    Worksheet worksheet = workbook.Worksheets["Grouping"];
                    workbook.Worksheets.ActiveWorksheet = worksheet;

                    #region #UngroupRows
                    // Ungroup four rows (from the third row to the sixth row) and display collapsed data.
                    worksheet.Rows.UnGroup(2, 5, true);

                    // Ungroup four rows (from the ninth row to the twelfth row).
                    worksheet.Rows.UnGroup(8, 11, false);

                    // Remove the outer group of rows.
                    worksheet.Rows.UnGroup(1, 12, false);
                    #endregion #UngroupRows
                }
                finally { workbook.EndUpdate(); }
            }

            static void UngroupColumnsValue(IWorkbook workbook)
            {
                workbook.LoadDocument("Documents\\SalesReport.xlsx");
                workbook.BeginUpdate();
                try
                {
                    Worksheet worksheet = workbook.Worksheets["Grouping"];
                    workbook.Worksheets.ActiveWorksheet = worksheet;

                    #region #UngroupColumns
                    // Ungroup four columns (from the column "C" to the column "F").
                    worksheet.Columns.UnGroup(2, 5, false);
                    #endregion #UngroupColumns
                }
                finally { workbook.EndUpdate(); }
            }

            static void AutoOutlineValue(IWorkbook workbook)
            {
                workbook.LoadDocument("Documents\\SalesReport.xlsx");
                workbook.BeginUpdate();
                try
                {
                    Worksheet worksheet = workbook.Worksheets["Sales Analysis"];
                    workbook.Worksheets.ActiveWorksheet = worksheet;

                    #region #AutoOutline
                    // Outline the data automatically based on the summary formulas.
                    worksheet.AutoOutline();
                    #endregion #AutoOutline
                }
                finally { workbook.EndUpdate(); }
            }

            static void SubtotalValue(IWorkbook workbook)
            {
                workbook.LoadDocument("Documents\\SalesReport.xlsx");
                workbook.BeginUpdate();
                try
                {
                    Worksheet worksheet = workbook.Worksheets["Regional Sales"];
                    workbook.Worksheets.ActiveWorksheet = worksheet;

                    #region #Subtotal
                    CellRange dataRange = worksheet["B3:E23"];
                    // Specify that subtotals should be calculated for the column "D". 
                    List<int> subtotalColumnsList = new List<int>();
                    subtotalColumnsList.Add(3);
                    // Insert subtotals by each change in the column "B" and calculate the SUM fuction for the related rows in the column "D".
                    worksheet.Subtotal(dataRange, 1, subtotalColumnsList, 9, "Total");
                    #endregion #Subtotal
                }
                finally { workbook.EndUpdate(); }
            }
    }
}
