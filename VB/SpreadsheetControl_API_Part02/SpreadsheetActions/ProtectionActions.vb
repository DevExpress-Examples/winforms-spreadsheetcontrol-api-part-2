Imports System
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraPrinting.Control
Imports System.Drawing

Namespace SpreadsheetControl_API
    Public NotInheritable Class ProtectionActions

#Region "Actions"
        Public Shared ProtectWorkbookAction As Action(Of IWorkbook) = AddressOf ProtectWorkbook
        Public Shared ProtectWorksheetAction As Action(Of IWorkbook) = AddressOf ProtectWorksheet
        Public Shared ProtectRangeAction As Action(Of IWorkbook) = AddressOf ProtectRange
#End Region

        Private Sub New()
        End Sub
        Private Shared Sub ProtectWorkbook(ByVal workbook As IWorkbook)
            '			#Region "#ProtectWorkbook"
            ' Protect workbook structure (prevents users from adding or deleting worksheets
            ' or from displaying hidden worksheets)
            workbook.BeginUpdate()
            If (Not workbook.IsProtected) Then
                workbook.Protect("password", True, False)
            End If
            workbook.Worksheets(0).Visible = False
            Dim worksheet As Worksheet = workbook.Worksheets(1)
            worksheet("D5").Value = "You are not allowed to add or delete a worksheet."
            worksheet("D6").Value = "Hidden worksheets cannot be displayed."
            workbook.EndUpdate()
            '			#End Region ' #ProtectWorkbook
        End Sub

        Private Shared Sub ProtectWorksheet(ByVal workbook As IWorkbook)
            '			#Region "#ProtectWorksheet"
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Protect a worksheet. Prevent users from making changes to worksheet elements.
            If (Not worksheet.IsProtected) Then
                worksheet.Protect("password", WorksheetProtectionPermissions.Default)
            End If
            workbook.BeginUpdate()
            worksheet("C3:F8").Borders.SetOutsideBorders(Color.Red, BorderLineStyle.Thin)
            worksheet("D5:E6").Merge()
            worksheet("D5").Value = "Try to change me!"
            worksheet("D5").Alignment.Vertical = SpreadsheetVerticalAlignment.Center
            workbook.EndUpdate()
            '			#End Region ' #ProtectWorksheet
        End Sub

        Private Shared Sub ProtectRange(ByVal workbook As IWorkbook)
            '			#Region "#ProtectRange"
            workbook.BeginUpdate()
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            worksheet("C3:E8").Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin)

            ' Give specific user permission to edit a range in a protected worksheet. 
            Dim protectedRange As ProtectedRange = worksheet.ProtectedRanges.Add("My Range", worksheet("C3:E8"))
            Dim permission As New EditRangePermission()
            permission.UserName = Environment.UserName
            permission.DomainName = Environment.UserDomainName
            permission.Deny = False
            protectedRange.SecurityDescriptor = protectedRange.CreateSecurityDescriptor(New EditRangePermission() {permission})
            protectedRange.SetPassword("123")

            ' Protect a worksheet.
            If (Not worksheet.IsProtected) Then
                worksheet.Protect("password", WorksheetProtectionPermissions.Default)
            End If

            worksheet.ActiveView.ShowGridlines = False
            workbook.EndUpdate()
            '			#End Region ' #ProtectRange
        End Sub
    End Class
End Namespace