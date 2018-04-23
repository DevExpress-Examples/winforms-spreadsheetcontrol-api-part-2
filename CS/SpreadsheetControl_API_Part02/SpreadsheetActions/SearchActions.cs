using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetControl_API
{
    class SearchActions
    {
        #region Actions
        public static Action<IWorkbook> SimpleSearchAction = SimpleSearchValue;
        public static Action<IWorkbook> AdvancedSearchAction = AdvancedSearchValue;
        #endregion

        static void SimpleSearchValue(IWorkbook workbook)
        {
            #region #SimpleSearch
            workbook.LoadDocument("Documents\\ExpenseReport.xlsx");
            workbook.Calculate();
            Worksheet worksheet = workbook.Worksheets[0];

            // Find and highlight cells containing the word "holiday".
            IEnumerable<Cell> searchResult = worksheet.Search("holiday");
            foreach (Cell cell in searchResult)
                cell.Fill.BackgroundColor = Color.LightGreen;
            #endregion #SimpleSearch

            // Add a note.
            worksheet["E1"].Value = "Find the word \"holiday\" in the expense report";
        }

        static void AdvancedSearchValue(IWorkbook workbook)
        {
            #region #AdvancedSearch
            workbook.LoadDocument("Documents\\ExpenseReport.xlsx");
            workbook.Calculate();
            Worksheet worksheet = workbook.Worksheets[0];

            // Specify the search term.
            string searchString = DateTime.Today.ToString("d");

            // Specify search options.
            SearchOptions options = new SearchOptions();
            options.SearchBy = SearchBy.Columns;
            options.SearchIn = SearchIn.Values;
            options.MatchEntireCellContents = true;

            // Find all cells containing today's date and paint them light-green.
            IEnumerable<Cell> searchResult = worksheet.Search(searchString, options);
            foreach (Cell cell in searchResult)
                cell.Fill.BackgroundColor = Color.LightGreen;
            #endregion #AdvancedSearch

            // Add a note.
            worksheet["E1"].Value = "Find today's date in the expense report";
        }
    }
}
