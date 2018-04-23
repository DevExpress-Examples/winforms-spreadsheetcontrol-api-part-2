Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports System.Diagnostics

Namespace SpreadsheetControl_API
	Partial Public Class Form1
		Inherits Form

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
'			#Region "GroupNodes"
			examples.Add(New SpreadsheetNode("Shapes"))
            examples.Add(New SpreadsheetNode("Custom Functions"))
            examples.Add(New SpreadsheetNode("Tables"))
'			#End Region

'			#Region "ExampleNodes"

			' Add nodes to the "Shapes" group of examples.
			examples(0).Groups.Add(New SpreadsheetExample("Insert a picture", ShapeActions.InsertShapeAction))
			examples(0).Groups.Add(New SpreadsheetExample("Insert a picture from URI", ShapeActions.InsertShapeFromUriAction))
			examples(0).Groups.Add(New SpreadsheetExample("Modify a picture", ShapeActions.ModifyShapeAction))

            ' Add nodes to the "Custom Functions" group of examples.
            examples(1).Groups.Add(New SpreadsheetExample("Add a SPHEREMASS function", CustomFunctionActions.SphereMassAction))

            ' Add nodes to the "Tables" group of examples.
            examples(2).Groups.Add(New SpreadsheetExample("Create a table", TableActions.CreateTableAction))
            examples(2).Groups.Add(New SpreadsheetExample("Apply a custom style", TableActions.CustomTableStyleAction))

'			#End Region
		End Sub

		Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
			treeList1.DataSource = examples
			treeList1.ExpandAll()
			treeList1.BestFitColumns()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
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
'			#Region "#LoadDocumentFromFile"
			' Load a workbook from the file.
			workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.OpenXml)
'			#End Region ' #LoadDocumentFromFile
		End Sub

		Private Sub SaveDocumentToFile()
'			#Region "#SaveDocumentToFile"
			' Save the modified document to the file.
			workbook.SaveDocument("Documents\SavedDocument.xlsx", DocumentFormat.OpenXml)
'			#End Region ' #SaveDocumentToFile
		End Sub
	End Class
End Namespace
