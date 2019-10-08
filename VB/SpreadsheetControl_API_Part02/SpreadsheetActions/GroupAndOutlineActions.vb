Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Drawing

Namespace SpreadsheetControl_API
    Friend Class GroupAndOutlineActions
        #Region "Actions"
        Public Shared GroupRowsAction As Action(Of IWorkbook) = AddressOf GroupRowsValue
        Public Shared GroupColumnsAction As Action(Of IWorkbook) = AddressOf GroupColumnsValue
        Public Shared UngroupRowsAction As Action(Of IWorkbook) = AddressOf UngroupRowsValue
        Public Shared UngroupColumnsAction As Action(Of IWorkbook) = AddressOf UngroupColumnsValue
        Public Shared AutoOutlineAction As Action(Of IWorkbook) = AddressOf AutoOutlineValue
        Public Shared SubtotalAction As Action(Of IWorkbook) = AddressOf SubtotalValue
        #End Region

            Private Shared Sub GroupRowsValue(ByVal workbook As IWorkbook)
                workbook.LoadDocument("Documents\SalesReport.xlsx")
                workbook.BeginUpdate()
                Try
                    Dim worksheet As Worksheet = workbook.Worksheets("Sales Analysis")
                    workbook.Worksheets.ActiveWorksheet = worksheet

'                    #Region "#GroupRows"
                    ' Group four rows starting from the third row and collapse the group.
                    worksheet.Rows.Group(2, 5, True)

                    ' Group four rows starting from the ninth row and expand the group.
                    worksheet.Rows.Group(8, 11, False)

                    ' Create the outer group of rows by grouping rows 2 through 13. 
                    worksheet.Rows.Group(1, 12, False)
'                    #End Region ' #GroupRows
                Finally
                    workbook.EndUpdate()
                End Try
            End Sub

            Private Shared Sub GroupColumnsValue(ByVal workbook As IWorkbook)
                workbook.LoadDocument("Documents\SalesReport.xlsx")
                workbook.BeginUpdate()
                Try
                    Dim worksheet As Worksheet = workbook.Worksheets("Sales Analysis")
                    workbook.Worksheets.ActiveWorksheet = worksheet

'                    #Region "#GroupColumns"
                    ' Group four columns starting from the third column "C" and expand the group.
                    worksheet.Columns.Group(2, 5, False)
'                    #End Region ' #GroupColumns
                Finally
                    workbook.EndUpdate()
                End Try
            End Sub

            Private Shared Sub UngroupRowsValue(ByVal workbook As IWorkbook)
                workbook.LoadDocument("Documents\SalesReport.xlsx")
                workbook.BeginUpdate()
                Try
                    Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
                    workbook.Worksheets.ActiveWorksheet = worksheet

'                    #Region "#UngroupRows"
                    ' Ungroup four rows (from the third row to the sixth row) and display collapsed data.
                    worksheet.Rows.UnGroup(2, 5, True)

                    ' Ungroup four rows (from the ninth row to the twelfth row).
                    worksheet.Rows.UnGroup(8, 11, False)

                    ' Remove the outer group of rows.
                    worksheet.Rows.UnGroup(1, 12, False)
'                    #End Region ' #UngroupRows
                Finally
                    workbook.EndUpdate()
                End Try
            End Sub

            Private Shared Sub UngroupColumnsValue(ByVal workbook As IWorkbook)
                workbook.LoadDocument("Documents\SalesReport.xlsx")
                workbook.BeginUpdate()
                Try
                    Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
                    workbook.Worksheets.ActiveWorksheet = worksheet

'                    #Region "#UngroupColumns"
                    ' Ungroup four columns (from the column "C" to the column "F").
                    worksheet.Columns.UnGroup(2, 5, False)
'                    #End Region ' #UngroupColumns
                Finally
                    workbook.EndUpdate()
                End Try
            End Sub

            Private Shared Sub AutoOutlineValue(ByVal workbook As IWorkbook)
                workbook.LoadDocument("Documents\SalesReport.xlsx")
                workbook.BeginUpdate()
                Try
                    Dim worksheet As Worksheet = workbook.Worksheets("Sales Analysis")
                    workbook.Worksheets.ActiveWorksheet = worksheet

'                    #Region "#AutoOutline"
                    ' Outline the data automatically based on the summary formulas.
                    worksheet.AutoOutline()
'                    #End Region ' #AutoOutline
                Finally
                    workbook.EndUpdate()
                End Try
            End Sub

            Private Shared Sub SubtotalValue(ByVal workbook As IWorkbook)
                workbook.LoadDocument("Documents\SalesReport.xlsx")
                workbook.BeginUpdate()
                Try
                    Dim worksheet As Worksheet = workbook.Worksheets("Regional Sales")
                    workbook.Worksheets.ActiveWorksheet = worksheet

                '                    #Region "#Subtotal"
                Dim dataRange As CellRange = worksheet("B3:E23")
                ' Specify that subtotals should be calculated for the column "D". 
                Dim subtotalColumnsList As New List(Of Integer)()
                    subtotalColumnsList.Add(3)
                    ' Insert subtotals by each change in the column "B" and calculate the SUM fuction for the related rows in the column "D".
                    worksheet.Subtotal(dataRange, 1, subtotalColumnsList, 9, "Total")
'                    #End Region ' #Subtotal
                Finally
                    workbook.EndUpdate()
                End Try
            End Sub
    End Class
End Namespace
