Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace SpreadsheetControl_API
    #Region "#samplecomparer"
    Friend Class SampleComparer
        Implements IComparer(Of DevExpress.Spreadsheet.CellValue)

        Public Function Compare(ByVal a As DevExpress.Spreadsheet.CellValue, ByVal b As DevExpress.Spreadsheet.CellValue) As Integer Implements IComparer(Of DevExpress.Spreadsheet.CellValue).Compare
            If (Not a.IsText) OrElse (Not b.IsText) Then
                Return 0
            End If
            If a.TextValue.Length = b.TextValue.Length Then
                Return 0
            End If
            Return If(a.TextValue.Length > b.TextValue.Length, 1, -1)
        End Function
    End Class
    #End Region ' #samplecomparer
End Namespace
