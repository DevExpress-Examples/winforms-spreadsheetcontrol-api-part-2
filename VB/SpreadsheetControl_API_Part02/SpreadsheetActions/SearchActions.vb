Imports Microsoft.VisualBasic
Imports System
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Drawing

Namespace SpreadsheetControl_API
    Public NotInheritable Class SearchActions
#Region "Actions"
        Public Shared SimpleSearchAction As Action(Of IWorkbook) = AddressOf SimpleSearchValue
        Public Shared AdvancedSearchAction As Action(Of IWorkbook) = AddressOf AdvancedSearchValue
#End Region


        Shared Sub SimpleSearchValue(ByVal workbook As IWorkbook)
            '            #Region "#Actions"
            workbook.LoadDocument("Documents\ExpenseReport.xlsx")
            workbook.Calculate()
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Find and highlight cells containing the word "holiday".
            Dim searchResult As IEnumerable(Of Cell) = worksheet.Search("holiday")
            For Each cell As Cell In searchResult
                cell.Fill.BackgroundColor = Color.LightGreen
            Next cell
            '			#End Region ' #SimpleSearch

            ' Add a note.
            worksheet("E1").Value = "Find the word ""holiday"" in the expense report"
        End Sub

        Shared Sub AdvancedSearchValue(ByVal workbook As IWorkbook)
            '			#Region "#AdvancedSearch"
            workbook.LoadDocument("Documents\ExpenseReport.xlsx")
            workbook.Calculate()
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Specify the search term.
            Dim searchString As String = DateTime.Today.ToString("d")

            ' Specify search options.
            Dim options As New SearchOptions()
            options.SearchBy = SearchBy.Columns
            options.SearchIn = SearchIn.Values
            options.MatchEntireCellContents = True

            ' Find all cells containing today's date and paint them light-green.
            Dim searchResult As IEnumerable(Of Cell) = worksheet.Search(searchString, options)
            For Each cell As Cell In searchResult
                cell.Fill.BackgroundColor = Color.LightGreen
            Next cell
            '			#End Region ' #AdvancedSearch
            ' Add a note.
            worksheet("E1").Value = "Find today's date in the expense report"
        End Sub
    End Class

End Namespace
