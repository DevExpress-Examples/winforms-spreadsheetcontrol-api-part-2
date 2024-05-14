using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetControl_API
{
    class AutoFilterActions
    {
        #region Actions
        public static Action<IWorkbook> ApplyFilterAction = ApplyFilter;
        public static Action<IWorkbook> FilterAndSortBySingleColumnAction = FilterAndSortBySingleColumn;
        public static Action<IWorkbook> FilterAndSortByMultipleColumnsAction = FilterAndSortByMultipleColumns;
        public static Action<IWorkbook> FilterNumericByConditionAction = FilterNumericByCondition;
        public static Action<IWorkbook> FilterTextByConditionAction = FilterTextByCondition;
        public static Action<IWorkbook> FilterDatesByConditionAction = FilterDatesByCondition;
        public static Action<IWorkbook> FilterByValuesAction = FilterByValue;
        public static Action<IWorkbook> FilterByMultipleValuesAction = FilterByMultipleValues;
        public static Action<IWorkbook> FilterMixedDataTypesByValuesAction = FilterMixedDataTypesByValues;
        public static Action<IWorkbook> Top10FilterAction = Top10FilterValue;
        public static Action<IWorkbook> DynamicFilterAction = DynamicFilterValue;
        public static Action<IWorkbook> FilterAndSortByColorAction = FilterAndSortByColorValue;
        public static Action<IWorkbook> FilterByBackgroundColorAction = FilterByBackgroundColorValue;
        public static Action<IWorkbook> FilterByFontColorAction = FilterByFontColorValue;
        public static Action<IWorkbook> FilterByFillColorAction = FilterByFillColorValue;
        public static Action<IWorkbook> ReapplyFilterAction = ReapplyFilterValue;
        public static Action<IWorkbook> ClearFilterAction = ClearFilter;
        public static Action<IWorkbook> DisableFilterAction = DisableFilter;
        #endregion

        static void ApplyFilter(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #ApplyFilter
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);
                #endregion #ApplyFilter
            }
            finally { workbook.EndUpdate(); }
        }

        static void FilterAndSortBySingleColumn(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #FilterSortBySingleColumn
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Sort the data in descending order by the first column.
                worksheet.AutoFilter.SortState.Sort(0, true);
                #endregion #FilterSortBySingleColumn
            }
            finally { workbook.EndUpdate(); }
        }

        static void FilterAndSortByMultipleColumns(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #FilterSortByMultipleColumns
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Sort the data in descending order by the first and third columns.
                List<SortCondition> sortConditions = new List<SortCondition>();
                sortConditions.Add(new SortCondition(0, true));
                sortConditions.Add(new SortCondition(2, true));
                worksheet.AutoFilter.SortState.Sort(sortConditions);
                #endregion #FilterSortByMultipleColumns
            }
            finally { workbook.EndUpdate(); }
        }

        static void FilterNumericByCondition(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #FilterByCondition
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter values in the "Sales" column that are in a range from 5000$ to 8000$.
                AutoFilterColumn sales = worksheet.AutoFilter.Columns[2];
                sales.ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThanOrEqual, 8000, FilterComparisonOperator.LessThanOrEqual, true);
                #endregion #FilterByCondition
            }
            finally { workbook.EndUpdate(); }
        }

        static void FilterTextByCondition(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #FilterTextByCondition
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter values in the "Product" column that contain "Gi" and include empty cells.
                AutoFilterColumn products = worksheet.AutoFilter.Columns[1];
                products.ApplyCustomFilter("*Gi*", FilterComparisonOperator.Equal, FilterValue.FilterByBlank, FilterComparisonOperator.Equal, false);
                #endregion #FilterTextByCondition
            }
            finally { workbook.EndUpdate(); }
        }

        static void FilterByValue(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #FilterByValue
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter the data in the "Product" column by a specific value.
                worksheet.AutoFilter.Columns[1].ApplyFilterCriteria("Mozzarella di Giovanni");
                #endregion #FilterByValue
            }
            finally { workbook.EndUpdate(); }
        }

        static void FilterByMultipleValues(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #FilterByValues
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter the data in the "Product" column by an array of values.
                worksheet.AutoFilter.Columns[1].ApplyFilterCriteria(new CellValue[] { "Mozzarella di Giovanni", "Gorgonzola Telino" });
                #endregion #FilterByValues
            }
            finally { workbook.EndUpdate(); }
        }

        static void FilterDatesByCondition(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #FilterDatesByCondition
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter values in the "Reported Date" column to display dates that are between June 1, 2014 and February 1, 2015.
                worksheet.AutoFilter.Columns[3].ApplyCustomFilter(new DateTime(2014, 6, 1), FilterComparisonOperator.GreaterThanOrEqual, new DateTime(2015, 2, 1), FilterComparisonOperator.LessThanOrEqual, true);
                #endregion #FilterDatesByCondition
            }
            finally { workbook.EndUpdate(); }
        }

        static void FilterMixedDataTypesByValues(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #FilterMixedDataTypesByValues
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);
                // Create date grouping item to filter January 2015 dates.
                IList<DateGrouping> groupings = new List<DateGrouping>();
                DateGrouping dateGroupingJan2015 = new DateGrouping(new DateTime(2015, 1, 1), DateTimeGroupingType.Month);
                groupings.Add(dateGroupingJan2015);

                // Filter the data in the "Reported Date" column to display values reported in January 2015.
                worksheet.AutoFilter.Columns[3].ApplyFilterCriteria("gennaio 2015", groupings);
                #endregion #FilterMixedDataTypesByValues
            }
            finally { workbook.EndUpdate(); }
        }

        static void Top10FilterValue(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #Top10Filter
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Apply a filter to the "Sales" column to display the top ten values.
                worksheet.AutoFilter.Columns[2].ApplyTop10Filter(Top10Type.Top10Items, 10);
                #endregion #Top10Filter
            }
            finally { workbook.EndUpdate(); }
        }

        static void DynamicFilterValue(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #DynamicFilter
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Apply a dynamic filter to the "Sales" column to display only values that are above the average.
                worksheet.AutoFilter.Columns[2].ApplyDynamicFilter(DynamicFilterType.AboveAverage);

                #endregion #DynamicFilter
            }
            finally { workbook.EndUpdate(); }
        }

        static void FilterAndSortByColorValue(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                #region #FilterAndSortByColor
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                // Enable filtering for the "B2:E23" cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Sort data in the "B2:E23" range
                // in descending order by column "D".
                Color color = worksheet["D12"].Font.Color;
                worksheet.AutoFilter.SortState.Sort(2, color, false);
                #endregion #FilterAndSortByColor
            }
            finally { workbook.EndUpdate(); }

        }

        static void FilterByBackgroundColorValue(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                #region #FilterByBackgroundColor
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                // Enable filtering for the "B2:E23" cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter values in the "Products" column by background color.
                AutoFilterColumn products = worksheet.AutoFilter.Columns[1];
                products.ApplyFillColorFilter(worksheet["C12"].FillColor);
                #endregion #FilterByBackgroundColor
            }
            finally { workbook.EndUpdate(); }

        }

        static void FilterByFillColorValue(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                #region #FilterByFillColor
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                // Enable filtering for the "B2:E23" cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter values in the "Products" column by fill color.
                AutoFilterColumn products = worksheet.AutoFilter.Columns[1];
                products.ApplyFillFilter(worksheet["C10"].Fill);
                #endregion #FilterByFillColor
            }
            finally { workbook.EndUpdate(); }


        }

        static void FilterByFontColorValue(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                #region #FilterByFontColor
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                // Enable filtering for the "B2:E23" cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter values in the "Sales" column by font color.
                AutoFilterColumn products = worksheet.AutoFilter.Columns[2];
                products.ApplyFontColorFilter(worksheet["D10"].Font.Color);
                #endregion #FilterByFontColor
            }
            finally { workbook.EndUpdate(); }


        }

        static void ReapplyFilterValue(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #ReapplyFilter
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter values in the "Sales" column that are greater than 5000$.
                worksheet.AutoFilter.Columns[2].ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan);

                // Change the data and reapply the filter.
                worksheet["D3"].Value = 5000;
                worksheet.AutoFilter.ReApply();
                #endregion #ReapplyFilter
            }
            finally { workbook.EndUpdate(); }
        }

        static void ClearFilter(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #ClearFilter
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Filter values in the "Sales" column that are greater than 5000$.
                worksheet.AutoFilter.Columns[2].ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan);

                // Clear the filter.
                worksheet.AutoFilter.Clear();
                #endregion #ClearFilter
            }
            finally { workbook.EndUpdate(); }
        }

        static void DisableFilter(IWorkbook workbook)
        {
            workbook.LoadDocument("Documents\\SalesReport.xlsx");
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["Regional sales"];
                workbook.Worksheets.ActiveWorksheet = worksheet;

                #region #DisableFilter
                // Enable filtering for the specified cell range.
                CellRange range = worksheet["B2:E23"];
                worksheet.AutoFilter.Apply(range);

                // Disable filtering for the entire worksheet.
                worksheet.AutoFilter.Disable();
                #endregion #DisableFilter
            }
            finally { workbook.EndUpdate(); }
        }

    }
}