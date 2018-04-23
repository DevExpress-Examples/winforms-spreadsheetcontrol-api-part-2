Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
#Region "#usings_CFunc"
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Functions
Imports DevExpress.XtraSpreadsheet
#End Region ' #usings_CFunc

Namespace SpreadsheetControl_API
	Public NotInheritable Class CustomFunctionActions

		Private Sub New()
		End Sub

		#Region "Actions"
		Public Shared SphereMassAction As Action(Of IWorkbook) = AddressOf SphereMassValue
		#End Region

		Private Shared Sub SphereMassValue(ByVal workbook As IWorkbook)
'			#Region "#customfunctionuse"
' Create a custom function and add it to the global scope.
Dim customFunction As New SphereMassFunction()
If Not workbook.CustomFunctions.Contains(customFunction.Name) Then
	workbook.CustomFunctions.Add(customFunction)
End If
'			#End Region ' #customfunctionuse

			workbook.BeginUpdate()
			Try
				Dim worksheet As Worksheet = workbook.Worksheets(0)
				worksheet.Range("A1:H1").ColumnWidthInCharacters = 12
				worksheet.Range("A1:H1").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center

				worksheet.DefinedNames.Add("seawater", "1025")
				worksheet.DefinedNames.Add("iron", "7870")
				worksheet.DefinedNames.Add("gold", "19300")

				worksheet("A1").Value = "Radius, m"
				worksheet("B1").Value = "Material"
				worksheet("C1").Value = "Mass, kg"
				worksheet("A2").Value = 0.1
				worksheet("B2").Value = ""
				worksheet("C2").FormulaInvariant = "=SPHEREMASS(A2)"
				worksheet("C2").NumberFormat = "#.##"
				worksheet("A3").Value = 0.1
				worksheet("B3").Value = "Seawater"
				worksheet("C3").FormulaInvariant = "=SPHEREMASS(A3, seawater)"
				worksheet("C3").NumberFormat = "#.##"
				worksheet("A4").Value = 0.1
				worksheet("B4").Value = "Iron"
				worksheet("C4").FormulaInvariant = "=SPHEREMASS(A4, iron)"
				worksheet("C4").NumberFormat = "#.##"
				worksheet("A5").Value = 0.1
				worksheet("B5").Value = "Gold"
				worksheet("C5").FormulaInvariant = "=SPHEREMASS(A5, gold)"
				worksheet("C5").NumberFormat = "#.##"
			Finally
				workbook.EndUpdate()
			End Try

		End Sub

	End Class


#Region "#customfunctiondef"
' Inheritance from Object is required for automatic VB.NET conversion
Public Class SphereMassFunction
	Inherits Object
	Implements ICustomFunction

	Private Const functionName As String = "SPHEREMASS"
	Private ReadOnly functionParameters() As ParameterInfo

	Public Sub New()
		' Missing optional parameters do not result in an error message.
		Me.functionParameters = New ParameterInfo() {
			New ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
			New ParameterInfo(ParameterType.Value, ParameterAttributes.Optional)
		}
	End Sub

	Public ReadOnly Property Name() As String Implements DevExpress.Spreadsheet.Functions.IFunction.Name
		Get
			Return functionName
		End Get
	End Property
	Private ReadOnly Property IFunction_Parameters() As ParameterInfo() Implements IFunction.Parameters
		Get
			Return functionParameters
		End Get
	End Property
	Private ReadOnly Property IFunction_ReturnType() As ParameterType Implements IFunction.ReturnType
		Get
			Return ParameterType.Value
		End Get
	End Property
	' Reevaluate cells on every recalculation.
	Private ReadOnly Property IFunction_Volatile() As Boolean Implements IFunction.Volatile
		Get
			Return True
		End Get
	End Property

	Private Function IFunction_Evaluate(ByVal parameters As IList(Of ParameterValue), ByVal context As EvaluationContext) As ParameterValue Implements IFunction.Evaluate
		Dim radius As Double
		Dim density As Double = 1000
		Dim radiusParameter As ParameterValue
		Dim densityParameter As ParameterValue

		If parameters.Count = 2 Then
			densityParameter = parameters(1)
			If densityParameter.IsError Then
				Return densityParameter
			Else
				density = densityParameter.NumericValue
			End If
		End If

		radiusParameter = parameters(0)
		If radiusParameter.IsError Then
			Return radiusParameter
		Else
			radius = radiusParameter.NumericValue
		End If

		Return (4 * Math.PI) / 3 * Math.Pow(radius,3) * density

	End Function
	Private Function IFunction_GetName(ByVal culture As CultureInfo) As String Implements IFunction.GetName
		Return functionName
	End Function
End Class
#End Region ' #customfunctiondef
End Namespace
