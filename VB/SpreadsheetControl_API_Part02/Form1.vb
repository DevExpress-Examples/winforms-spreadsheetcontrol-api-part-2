Imports System
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports System.Diagnostics

Namespace SpreadsheetControl_API

    Public Partial Class Form1
        Inherits Form

        Private workbook As IWorkbook

        Public Sub New()
            InitializeComponent()
            ' Access a workbook.
            workbook = spreadsheetControl1.Document
            InitTreeListControl()
        End Sub

        Private Sub InitTreeListControl()
            Dim examples As GroupsOfSpreadsheetExamples = New GroupsOfSpreadsheetExamples()
            InitData(examples)
            DataBinding(examples)
        End Sub

        Private Sub InitData(ByVal examples As GroupsOfSpreadsheetExamples)
'#Region "GroupNodes"
            examples.Add(New SpreadsheetNode("Shapes"))
            examples.Add(New SpreadsheetNode("Custom Functions"))
'#End Region
'#Region "ExampleNodes"
            ' Add nodes to the "Shapes" group of examples.
            examples(0).Groups.Add(New SpreadsheetExample("Insert a picture", InsertShapeAction))
            examples(0).Groups.Add(New SpreadsheetExample("Insert a picture from URI", InsertShapeFromUriAction))
            examples(0).Groups.Add(New SpreadsheetExample("Modify a picture", ModifyShapeAction))
            ' Add nodes to the "Cuistom Function" group of examples.
            examples(1).Groups.Add(New SpreadsheetExample("Add a SPHEREMASS function", SphereMassAction))
'#End Region
        End Sub

        Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
            treeList1.DataSource = examples
            treeList1.ExpandAll()
            treeList1.BestFitColumns()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            LoadDocumentFromFile()
            Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
            If example Is Nothing Then Return
            Dim action As Action(Of IWorkbook) = example.Action
            action(workbook)
            spreadsheetControl1.Refresh()
            SaveDocumentToFile()
        End Sub

        ' ------------------- Load and Save a Document -------------------
        Private Sub LoadDocumentFromFile()
'#Region "#LoadDocumentFromFile"
            ' Load a workbook from the file.
            workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.OpenXml)
'#End Region  ' #LoadDocumentFromFile
        End Sub

        Private Sub SaveDocumentToFile()
'#Region "#SaveDocumentToFile"
            ' Save the modified document to the file.
            workbook.SaveDocument("Documents\SavedDocument.xlsx", DocumentFormat.OpenXml)
'#End Region  ' #SaveDocumentToFile
        End Sub
    End Class
End Namespace
