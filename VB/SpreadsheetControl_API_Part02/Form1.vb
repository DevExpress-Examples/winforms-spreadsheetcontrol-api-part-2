Imports System
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports System.Diagnostics

Namespace SpreadsheetControl_API
    Partial Public Class Form1
        Inherits DevExpress.XtraEditors.XtraForm

        Private workbook As IWorkbook

        Public Sub New()
            InitializeComponent()

            ' Access a workbook.
            workbook = spreadsheetControl1.Document

            InitTreeListControl()

        End Sub

        Private Sub InitTreeListControl()
            Dim examples As New GroupsOfSpreadsheetExamples()
            InitData(examples)
            DataBinding(examples)
        End Sub

        Private Sub InitData(ByVal examples As GroupsOfSpreadsheetExamples)
'            #Region "GroupNodes"
            examples.Add(New SpreadsheetNode("Pictures"))
            examples.Add(New SpreadsheetNode("Custom Functions"))
            examples.Add(New SpreadsheetNode("Tables"))
            examples.Add(New SpreadsheetNode("Protection"))
            examples.Add(New SpreadsheetNode("Sort"))
            examples.Add(New SpreadsheetNode("Search"))
            examples.Add(New SpreadsheetNode("Export"))
            examples.Add(New SpreadsheetNode("Group Data"))
            examples.Add(New SpreadsheetNode("Filter Data"))
            examples.Add(New SpreadsheetNode("Document Properties"))
'            #End Region

'            #Region "ExampleNodes"

            ' Add nodes to the "Pictures" group of examples.
            examples(0).Groups.Add(New SpreadsheetExample("Insert picture", PictureActions.InsertPictureAction))
            examples(0).Groups.Add(New SpreadsheetExample("Insert picture from URI", PictureActions.InsertPictureFromUriAction))
            examples(0).Groups.Add(New SpreadsheetExample("Move picture", PictureActions.MovePictureAction))
            examples(0).Groups.Add(New SpreadsheetExample("Rotate picture", PictureActions.RotatePictureAction))
            examples(0).Groups.Add(New SpreadsheetExample("Bring picture to front", PictureActions.ChangeZOrderAction))
            examples(0).Groups.Add(New SpreadsheetExample("Add hyperlink", PictureActions.InsertHyperlinkAction))


            ' Add nodes to the "Custom Functions" group of examples.
            examples(1).Groups.Add(New SpreadsheetExample("Add UDF(user defined function)", CustomFunctionActions.SphereMassAction))

            ' Add nodes to the "Tables" group of examples.
            examples(2).Groups.Add(New SpreadsheetExample("Create table", TableActions.CreateTableAction))
            examples(2).Groups.Add(New SpreadsheetExample("Apply custom style", TableActions.CustomTableStyleAction))

            ' Add nodes to the "Protection" group of examples.
            examples(3).Groups.Add(New SpreadsheetExample("Protect workbook", ProtectionActions.ProtectWorkbookAction))
            examples(3).Groups.Add(New SpreadsheetExample("Protect worksheet", ProtectionActions.ProtectWorksheetAction))
            examples(3).Groups.Add(New SpreadsheetExample("Protect range", ProtectionActions.ProtectRangeAction))

            ' Add nodes to the "Sort" group of examples.
            examples(4).Groups.Add(New SpreadsheetExample("Simple sort", SortActions.SimpleSortAction))
            examples(4).Groups.Add(New SpreadsheetExample("Sort in descending order", SortActions.DescendingOrderAction))
            examples(4).Groups.Add(New SpreadsheetExample("Sort using custom comparer", SortActions.SelectComparerAction))
            examples(4).Groups.Add(New SpreadsheetExample("Sort by column", SortActions.SortBySpecifiedColumnAction))
            examples(4).Groups.Add(New SpreadsheetExample("Sort by multiple columns", SortActions.SortByMultipleColumnsAction))

            ' Add nodes to the "Search" group of examples.
            examples(5).Groups.Add(New SpreadsheetExample("Simple search", SearchActions.SimpleSearchAction))
            examples(5).Groups.Add(New SpreadsheetExample("Search with options", SearchActions.AdvancedSearchAction))

            ' Add nodes to the "Export" group of examples.
            examples(6).Groups.Add(New SpreadsheetExample("Export to HTML", ExportActions.ExportToHTMLAction))

            ' Add nodes to the "Group Data" group of examples.
            examples(7).Groups.Add(New SpreadsheetExample("Group Rows", GroupAndOutlineActions.GroupRowsAction))
            examples(7).Groups.Add(New SpreadsheetExample("Group Columns", GroupAndOutlineActions.GroupColumnsAction))
            examples(7).Groups.Add(New SpreadsheetExample("Unroup Rows", GroupAndOutlineActions.UngroupRowsAction))
            examples(7).Groups.Add(New SpreadsheetExample("Unroup Columns", GroupAndOutlineActions.UngroupColumnsAction))
            examples(7).Groups.Add(New SpreadsheetExample("Create an Auto Outline", GroupAndOutlineActions.AutoOutlineAction))
            examples(7).Groups.Add(New SpreadsheetExample("Insert Subtotals", GroupAndOutlineActions.SubtotalAction))

            ' Add nodes to the "Filter Data" group of examples.
            examples(8).Groups.Add(New SpreadsheetExample("Enable filtering", AutoFilterActions.ApplyFilterAction))
            examples(8).Groups.Add(New SpreadsheetExample("Sort by single column", AutoFilterActions.FilterAndSortBySingleColumnAction))
            examples(8).Groups.Add(New SpreadsheetExample("Sort by multiple columns", AutoFilterActions.FilterAndSortByMultipleColumnsAction))
            examples(8).Groups.Add(New SpreadsheetExample("Custom number filter", AutoFilterActions.FilterNumericByConditionAction))
            examples(8).Groups.Add(New SpreadsheetExample("Custom text filter", AutoFilterActions.FilterTextByConditionAction))
            examples(8).Groups.Add(New SpreadsheetExample("Custom date filter", AutoFilterActions.FilterDatesByConditionAction))
            examples(8).Groups.Add(New SpreadsheetExample("Filter by single value", AutoFilterActions.FilterByValuesAction))
            examples(8).Groups.Add(New SpreadsheetExample("Filter by multiple values", AutoFilterActions.FilterByMultipleValuesAction))
            examples(8).Groups.Add(New SpreadsheetExample("Filter mixed data types by values", AutoFilterActions.FilterMixedDataTypesByValuesAction))
            examples(8).Groups.Add(New SpreadsheetExample("Apply Top 10 filter", AutoFilterActions.Top10FilterAction))
            examples(8).Groups.Add(New SpreadsheetExample("Apply dynamic filter", AutoFilterActions.DynamicFilterAction))
            examples(8).Groups.Add(New SpreadsheetExample("Reapply filter", AutoFilterActions.ReapplyFilterAction))
            examples(8).Groups.Add(New SpreadsheetExample("Clear filter", AutoFilterActions.ClearFilterAction))
            examples(8).Groups.Add(New SpreadsheetExample("Disable filtering", AutoFilterActions.DisableFilterAction))

            ' Add nodes to the "Document Properties" group of examples.
            examples(9).Groups.Add(New SpreadsheetExample("Built-in properties", DocumentPropertiesActions.BuiltInPropertiesAction))
            examples(9).Groups.Add(New SpreadsheetExample("Custom properties", DocumentPropertiesActions.CustomPropertiesAction))

'            #End Region
        End Sub

        Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
            treeList1.DataSource = examples
            treeList1.ExpandAll()
            treeList1.BestFitColumns()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            workbook.Options.Culture = System.Globalization.CultureInfo.CurrentCulture
            LoadDocumentFromFile()
            Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
            If example Is Nothing Then
                Return
            End If
            Dim action As Action(Of IWorkbook) = example.Action
            action(workbook)
            Me.spreadsheetControl1.Refresh()
            SaveDocumentToFile()
        End Sub

        ' ------------------- Load and Save a Document -------------------
        Private Sub LoadDocumentFromFile()
'            #Region "#LoadDocumentFromFile"
            ' Load a workbook from the file.
            workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.OpenXml)
'            #End Region ' #LoadDocumentFromFile
        End Sub

        Private Sub SaveDocumentToFile()
'            #Region "#SaveDocumentToFile"
            ' Save the modified document to the file.
            workbook.SaveDocument("Documents\SavedDocument.xlsx", DocumentFormat.OpenXml)
'            #End Region ' #SaveDocumentToFile
        End Sub
    End Class
End Namespace
