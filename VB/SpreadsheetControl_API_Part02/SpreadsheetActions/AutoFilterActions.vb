Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms

Namespace SpreadsheetControl_API
    Friend Class AutoFilterActions
        #Region "Actions"
        Public Shared ApplyFilterAction As Action(Of IWorkbook) = AddressOf ApplyFilter
        Public Shared FilterAndSortBySingleColumnAction As Action(Of IWorkbook) = AddressOf FilterAndSortBySingleColumn
        Public Shared FilterAndSortByMultipleColumnsAction As Action(Of IWorkbook) = AddressOf FilterAndSortByMultipleColumns
        Public Shared FilterNumericByConditionAction As Action(Of IWorkbook) = AddressOf FilterNumericByCondition
        Public Shared FilterTextByConditionAction As Action(Of IWorkbook) = AddressOf FilterTextByCondition
        Public Shared FilterDatesByConditionAction As Action(Of IWorkbook) = AddressOf FilterDatesByCondition
        Public Shared FilterByValuesAction As Action(Of IWorkbook) = AddressOf FilterByValue
        Public Shared FilterByMultipleValuesAction As Action(Of IWorkbook) = AddressOf FilterByMultipleValues
        Public Shared FilterMixedDataTypesByValuesAction As Action(Of IWorkbook) = AddressOf FilterMixedDataTypesByValues
        Public Shared Top10FilterAction As Action(Of IWorkbook) = AddressOf Top10FilterValue
        Public Shared DynamicFilterAction As Action(Of IWorkbook) = AddressOf DynamicFilterValue
        Public Shared ReapplyFilterAction As Action(Of IWorkbook) = AddressOf ReapplyFilterValue
        Public Shared ClearFilterAction As Action(Of IWorkbook) = AddressOf ClearFilter
        Public Shared DisableFilterAction As Action(Of IWorkbook) = AddressOf DisableFilter
        #End Region

        Private Shared Sub ApplyFilter(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#ApplyFilter"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)
'                #End Region ' #ApplyFilter
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub FilterAndSortBySingleColumn(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#FilterSortBySingleColumn"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Sort the data in descending order by the first column.
                worksheet.AutoFilter.SortState.Sort(0, True)
'                #End Region ' #FilterSortBySingleColumn
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub FilterAndSortByMultipleColumns(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#FilterSortByMultipleColumns"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Sort the data in descending order by the first and third columns.
                Dim sortConditions As New List(Of SortCondition)()
                sortConditions.Add(New SortCondition(0, True))
                sortConditions.Add(New SortCondition(2, True))
                worksheet.AutoFilter.SortState.Sort(sortConditions)
'                #End Region ' #FilterSortByMultipleColumns
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub FilterNumericByCondition(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#FilterByCondition"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Filter values in the "Sales" column that are in a range from 5000$ to 8000$.
                Dim sales As AutoFilterColumn = worksheet.AutoFilter.Columns(2)
                sales.ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThanOrEqual, 8000, FilterComparisonOperator.LessThanOrEqual, True)
'                #End Region ' #FilterByCondition
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub FilterTextByCondition(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#FilterTextByCondition"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Filter values in the "Product" column that contain "Gi" and include empty cells.
                Dim products As AutoFilterColumn = worksheet.AutoFilter.Columns(1)
                products.ApplyCustomFilter("*Gi*", FilterComparisonOperator.Equal, FilterValue.FilterByBlank, FilterComparisonOperator.Equal, False)
'                #End Region ' #FilterTextByCondition
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub FilterByValue(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#FilterByValue"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Filter the data in the "Product" column by a specific value.
                worksheet.AutoFilter.Columns(1).ApplyFilterCriteria("Mozzarella di Giovanni")
'                #End Region ' #FilterByValue
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub FilterByMultipleValues(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#FilterByValues"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Filter the data in the "Product" column by an array of values.
                worksheet.AutoFilter.Columns(1).ApplyFilterCriteria(New CellValue() { "Mozzarella di Giovanni", "Gorgonzola Telino"})
'                #End Region ' #FilterByValues
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub FilterDatesByCondition(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '            #Region "#FilterDatesByCondition"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

            ' Filter values in the "Reported Date" column to display dates that are between June 1, 2014 and February 1, 2015.
            worksheet.AutoFilter.Columns(3).ApplyCustomFilter(New Date(2014, 6, 1), FilterComparisonOperator.GreaterThanOrEqual, New Date(2015, 2, 1), FilterComparisonOperator.LessThanOrEqual, True)
'            #End Region ' #FilterDatesByCondition
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub FilterMixedDataTypesByValues(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#FilterMixedDataTypesByValues"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)
                ' Create date grouping item to filter January 2015 dates.
                Dim groupings As IList(Of DateGrouping) = New List(Of DateGrouping)()
                Dim dateGroupingJan2015 As New DateGrouping(New Date(2015, 1, 1), DateTimeGroupingType.Month)
                groupings.Add(dateGroupingJan2015)

                ' Filter the data in the "Reported Date" column to display values reported in January 2015.
                worksheet.AutoFilter.Columns(3).ApplyFilterCriteria("gennaio 2015", groupings)
'                #End Region ' #FilterMixedDataTypesByValues
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub Top10FilterValue(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#Top10Filter"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Apply a filter to the "Sales" column to display the top ten values.
                worksheet.AutoFilter.Columns(2).ApplyTop10Filter(Top10Type.Top10Items, 10)
'                #End Region ' #Top10Filter
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub DynamicFilterValue(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#DynamicFilter"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Apply a dynamic filter to the "Sales" column to display only values that are above the average.
                worksheet.AutoFilter.Columns(2).ApplyDynamicFilter(DynamicFilterType.AboveAverage)

                '                #End Region ' #DynamicFilter
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub ReapplyFilterValue(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#ReapplyFilter"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Filter values in the "Sales" column that are greater than 5000$.
                worksheet.AutoFilter.Columns(2).ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan)

                ' Change the data and reapply the filter.
                worksheet("D3").Value = 5000
                worksheet.AutoFilter.ReApply()
'                #End Region ' #ReapplyFilter
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub ClearFilter(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#ClearFilter"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Filter values in the "Sales" column that are greater than 5000$.
                worksheet.AutoFilter.Columns(2).ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan)

                ' Clear the filter.
                worksheet.AutoFilter.Clear()
'                #End Region ' #ClearFilter
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub DisableFilter(ByVal workbook As IWorkbook)
            workbook.LoadDocument("Documents\SalesReport.xlsx")
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
                workbook.Worksheets.ActiveWorksheet = worksheet

                '                #Region "#DisableFilter"
                ' Enable filtering for the specified cell range.
                Dim range As CellRange = worksheet("B2:E23")
                worksheet.AutoFilter.Apply(range)

                ' Disable filtering for the entire worksheet.
                worksheet.AutoFilter.Disable()
'                #End Region ' #DisableFilter
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

    End Class
End Namespace