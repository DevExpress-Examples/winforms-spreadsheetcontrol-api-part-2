using System;
using DevExpress.Spreadsheet;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrinting.Control;
using System.Drawing;

namespace SpreadsheetControl_API
{
    public static class ProtectionActions {

        #region Actions
        public static Action<IWorkbook> ProtectWorkbookAction = ProtectWorkbook;
        public static Action<IWorkbook> ProtectWorksheetAction = ProtectWorksheet;
        public static Action<IWorkbook> ProtectRangeAction = ProtectRange;

        #endregion
        
        static void ProtectWorkbook(IWorkbook workbook) {
            #region #ProtectWorkbook
            // Protect workbook structure (prevents users from adding or deleting worksheets
            // or from displaying hidden worksheets).
            workbook.BeginUpdate();
            if (!workbook.IsProtected)
                workbook.Protect("password", true, false);
            workbook.Worksheets[0].Visible = false;
            Worksheet worksheet = workbook.Worksheets[1];
            worksheet["D5"].Value = "You are not allowed to add or delete a worksheet.";
            worksheet["D6"].Value = "Hidden worksheets cannot be displayed.";
            workbook.EndUpdate();
            #endregion #ProtectWorkbook
        }

        static void ProtectWorksheet(IWorkbook workbook) {
            #region #ProtectWorksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Protect a worksheet. Prevent users from making changes to worksheet elements.
            if (!worksheet.IsProtected)
                worksheet.Protect("password", WorksheetProtectionPermissions.Default);
            workbook.BeginUpdate();
            worksheet["C3:F8"].Borders.SetOutsideBorders(Color.Red, BorderLineStyle.Thin);
            worksheet["D5:E6"].Merge();
            worksheet["D5"].Value = "Try to change me!";
            worksheet["D5"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            workbook.EndUpdate();
            #endregion #ProtectWorksheet
        }

        static void ProtectRange(IWorkbook workbook) {
            #region #ProtectRange
workbook.BeginUpdate();
Worksheet worksheet = workbook.Worksheets[0];
worksheet["C3:E8"].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);

// Give specific user permission to edit a range in a protected worksheet. 
ProtectedRange protectedRange = worksheet.ProtectedRanges.Add("My Range", worksheet["C3:E8"]);
EditRangePermission permission = new EditRangePermission();
permission.UserName = Environment.UserName;
permission.DomainName = Environment.UserDomainName;
permission.Deny = false;
protectedRange.SecurityDescriptor = protectedRange.CreateSecurityDescriptor(new EditRangePermission[] { permission });
protectedRange.SetPassword("123");

// Protect a worksheet.
if (!worksheet.IsProtected)
    worksheet.Protect("password", WorksheetProtectionPermissions.Default);
            
worksheet.ActiveView.ShowGridlines = false;
workbook.EndUpdate();
            #endregion #ProtectRange
        }
    }
}
