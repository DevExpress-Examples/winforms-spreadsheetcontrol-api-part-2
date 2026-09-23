Imports DevExpress.Spreadsheet
Imports System
Imports System.Drawing

Namespace SpreadsheetControl_API

    Public Module TableActions

'#Region "Actions"
        Public CreateTableAction As Action(Of IWorkbook) = AddressOf CreateTable

        Public CustomTableStyleAction As Action(Of IWorkbook) = AddressOf CustomTableStyle

'#End Region
        Private Sub CreateTable(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            GenerateTableData(worksheet)
'#Region "#CreateTable"
            ' Insert a table in the worksheet.
            Dim table As Table = worksheet.Tables.Add(worksheet("B2:F5"), True)
            ' Format the table by applying a built-in table style.
            table.Style = workbook.TableStyles(BuiltInTableStyleId.TableStyleMedium27)
            ' Access table columns and name them.
            Dim productColumn As TableColumn = table.Columns(0)
            productColumn.Name = "Product"
            Dim priceColumn As TableColumn = table.Columns(1)
            priceColumn.Name = "Price"
            Dim quantityColumn As TableColumn = table.Columns(2)
            quantityColumn.Name = "Quantity"
            Dim discountColumn As TableColumn = table.Columns(3)
            discountColumn.Name = "Discount"
            Dim amountColumn As TableColumn = table.Columns(4)
            amountColumn.Name = "Amount"
            ' Set the formula to calculate the amount per product 
            ' and display results in the "Amount" column.
            amountColumn.Formula = "=[Price]*[Quantity]*(1-[Discount])"
            ' Display the total row in the table.
            table.ShowTotals = True
            ' Set the label and function to display the sum of the "Amount" column.
            discountColumn.TotalRowLabel = "Total:"
            amountColumn.TotalRowFunction = TotalRowFunction.Sum
            ' Specify the number format for each column.
            priceColumn.DataRange.NumberFormat = "$#,##0.00"
            discountColumn.DataRange.NumberFormat = "0.0%"
            amountColumn.Range.NumberFormat = "$#,##0.00;$#,##0.00;"""";@"
            ' Specify horizontal alignment for header and total rows of the table.
            table.HeaderRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            table.TotalRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            ' Specify horizontal alignment to display data in all columns except the first one.
            For i As Integer = 1 To table.Columns.Count - 1
                table.Columns(i).DataRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            Next

            ' Set the width of table columns.
            table.Range.ColumnWidthInCharacters = 10
'#End Region  ' #CreateTable
            workbook.EndUpdate()
        End Sub

        Private Sub CustomTableStyle(ByVal workbook As IWorkbook)
            CreateTable(workbook)
            workbook.BeginUpdate()
            Dim worksheet As Worksheet = workbook.Worksheets(0)
'#Region "#CustomTableStyle"
            ' Access a table.
            Dim table As Table = worksheet.Tables(0)
            Dim styleName As String = "testTableStyle"
            ' If the style under the specified name already exists in the collection,
            If workbook.TableStyles.Contains(styleName) Then
                ' apply this style to the table.
                table.Style = workbook.TableStyles(styleName)
            Else
                ' Add a new table style under the "testTableStyle" name to the TableStyles collection.
                Dim lCustomTableStyle As TableStyle = workbook.TableStyles.Add("testTableStyle")
                ' Modify the required formatting characteristics of the table style. 
                ' Specify the format for different table elements.
                lCustomTableStyle.BeginUpdate()
                Try
                    lCustomTableStyle.TableStyleElements(TableStyleElementType.WholeTable).Font.Color = Color.FromArgb(107, 107, 107)
                    ' Specify formatting characteristics for the table header row. 
                    Dim headerRowStyle As TableStyleElement = lCustomTableStyle.TableStyleElements(TableStyleElementType.HeaderRow)
                    headerRowStyle.Fill.BackgroundColor = Color.FromArgb(64, 66, 166)
                    headerRowStyle.Font.Color = Color.White
                    headerRowStyle.Font.Bold = True
                    ' Specify formatting characteristics for the table total row. 
                    Dim totalRowStyle As TableStyleElement = lCustomTableStyle.TableStyleElements(TableStyleElementType.TotalRow)
                    totalRowStyle.Fill.BackgroundColor = Color.FromArgb(115, 193, 211)
                    totalRowStyle.Font.Color = Color.White
                    totalRowStyle.Font.Bold = True
                    ' Specify banded row formatting for the table.
                    Dim secondRowStripeStyle As TableStyleElement = lCustomTableStyle.TableStyleElements(TableStyleElementType.SecondRowStripe)
                    secondRowStripeStyle.Fill.BackgroundColor = Color.FromArgb(234, 234, 234)
                    secondRowStripeStyle.StripeSize = 1
                Finally
                    lCustomTableStyle.EndUpdate()
                End Try

                ' Apply the created custom style to the table.
                table.Style = lCustomTableStyle
            End If

'#End Region  ' #CustomTableStyle
            workbook.EndUpdate()
        End Sub

        Public Sub GenerateTableData(ByVal sheet As Worksheet)
            sheet.Cells("B3").SetValue("Chocolade")
            sheet.Cells("B4").SetValue("Konbu")
            sheet.Cells("B5").SetValue("Geitost")
            sheet.Cells("C3").SetValue(5.0)
            sheet.Cells("C4").SetValue(9.0)
            sheet.Cells("C5").SetValue(15.0)
            sheet.Cells("D3").SetValue(15)
            sheet.Cells("D4").SetValue(55)
            sheet.Cells("D5").SetValue(70)
            sheet.Cells("E3").SetValue(0.03)
            sheet.Cells("E4").SetValue(0.1)
            sheet.Cells("E5").SetValue(0.07)
        End Sub
    End Module
End Namespace
