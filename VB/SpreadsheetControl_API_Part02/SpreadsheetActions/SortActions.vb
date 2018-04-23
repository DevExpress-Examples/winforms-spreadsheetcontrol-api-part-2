Imports System
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports System.Globalization
Imports System.Collections.Generic

Namespace SpreadsheetControl_API
    Public NotInheritable Class SortActions

        Private Sub New()
        End Sub

        #Region "Actions"
        Public Shared SimpleSortAction As Action(Of IWorkbook) = AddressOf SimpleSortValue
        Public Shared DescendingOrderAction As Action(Of IWorkbook) = AddressOf DescendingOrderValue
        Public Shared SelectComparerAction As Action(Of IWorkbook) = AddressOf SelectComparerValue
        Public Shared SortBySpecifiedColumnAction As Action(Of IWorkbook) = AddressOf SortBySpecifiedColumnValue
        Public Shared SortByMultipleColumnsAction As Action(Of IWorkbook) = AddressOf SortByMultipleColumnsValue
        #End Region

        Private Shared Sub SimpleSortValue(ByVal workbook As IWorkbook)
'            #Region "#SimpleSort"
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Fill in the range.
            worksheet.Cells("A2").Value = "Donald Dozier Bradley"
            worksheet.Cells("A3").Value = "Tony Charles Mccallum-Geteer"
            worksheet.Cells("A4").Value = "Calvin Liu"
            worksheet.Cells("A5").Value = "Anita A Boyd"
            worksheet.Cells("A6").Value = "Angela R. Scott"
            worksheet.Cells("A7").Value = "D Fox"

            ' Sort the range in ascending order.
            Dim range As Range = worksheet.Range("A2:A7")
            worksheet.Sort(range)

            ' Create a heading.
            Dim header As Range = worksheet.Range("A1")
            header(0).Value = "Ascending order"
            header.ColumnWidthInCharacters = 30
            header.Style = workbook.Styles("Heading 1")
'            #End Region ' #SimpleSort
        End Sub

        Private Shared Sub DescendingOrderValue(ByVal workbook As IWorkbook)
'            #Region "#DescendingOrder"
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Fill in the range.
            worksheet.Cells("A2").Value = "Donald Dozier Bradley"
            worksheet.Cells("A3").Value = "Tony Charles Mccallum-Geteer"
            worksheet.Cells("A4").Value = "Calvin Liu"
            worksheet.Cells("A5").Value = "Anita A Boyd"
            worksheet.Cells("A6").Value = "Angela R. Scott"
            worksheet.Cells("A7").Value = "D Fox"

            ' Sort the range in descending order.
            Dim range As Range = worksheet.Range("A2:A7")
            worksheet.Sort(range, False)

            ' Create a heading.
            Dim header As Range = worksheet.Range("A1")
            header(0).Value = "Descending order"
            header.ColumnWidthInCharacters = 30
            header.Style = workbook.Styles("Heading 1")
'            #End Region ' #DescendingOrder
        End Sub

        Private Shared Sub SelectComparerValue(ByVal workbook As IWorkbook)
'            #Region "#SelectComparer"
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Fill in the range.
            worksheet.Cells("A2").Value = "Donald Dozier Bradley"
            worksheet.Cells("A3").Value = "Tony Charles Mccallum-Geteer"
            worksheet.Cells("A4").Value = "Calvin Liu"
            worksheet.Cells("A5").Value = "Anita A Boyd"
            worksheet.Cells("A6").Value = "Angela R. Scott"
            worksheet.Cells("A7").Value = "D Fox"

            ' Sort values using a custom comparer.
            Dim range As Range = worksheet.Range("A2:A7")
            worksheet.Sort(range, 0, New SampleComparer())

            ' Create a heading.
            Dim header As Range = worksheet.Range("A1")
            header(0).Value = "Use a custom comparer"
            header.ColumnWidthInCharacters = 30
            header.Style = workbook.Styles("Heading 1")
'            #End Region ' #SelectComparer
        End Sub

        Private Shared Sub SortBySpecifiedColumnValue(ByVal workbook As IWorkbook)
'            #Region "#SortBySpecifiedColumn"
            workbook.LoadDocument("Documents\Sortsample.xlsx")
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Sort by a column with offset = 6 in the range being sorted.
            ' Use ascending order.
            Dim range As Range = worksheet.Range("A3:F22")
            worksheet.Sort(range, 3)

            ' Add a note.
            worksheet("D1").Value = "Sort by column with index = 3 in ascending order"
            worksheet.Visible = True
'            #End Region ' #SortBySpecifiedColumn
        End Sub

        Private Shared Sub SortByMultipleColumnsValue(ByVal workbook As IWorkbook)
'            #Region "#SortByMultipleColumns"
            workbook.LoadDocument("Documents\Sortsample.xlsx")
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Create sorting fields.
            Dim fields As New List(Of SortField)()

            ' First sorting field. First column (offset = 0) will be sorted using ascending order.
            Dim sortField1 As New SortField()
            sortField1.ColumnOffset = 0
            sortField1.Comparer = worksheet.Comparers.Ascending
            fields.Add(sortField1)

            ' Second sorting field. Second column (offset = 1) will be sorted using ascending order.
            Dim sortField2 As New SortField()
            sortField2.ColumnOffset = 1
            sortField2.Comparer = worksheet.Comparers.Ascending
            fields.Add(sortField2)

            ' Sort the range by sorting fields.
            Dim range As Range = worksheet.Range("A3:F22")
            worksheet.Sort(range, fields)

'            #End Region ' #SortByMultipleColumns
            ' Add a note.
            worksheet("D1").Value = "Sort by two columns: first and second in ascending order"
        End Sub
    End Class
End Namespace
