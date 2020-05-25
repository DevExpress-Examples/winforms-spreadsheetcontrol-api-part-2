using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet.Export;
using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SpreadsheetControl_API
{
    class ExportActions
    {
        #region Actions
        public static Action<IWorkbook> ExportToHTMLAction = ExportToHTMLValue;
        #endregion

        private static void ExportToHTMLValue(IWorkbook workbook)
        {
            #region #ExportToHTML
            workbook.LoadDocument("Documents\\ExpenseReport.xlsx");
            workbook.Calculate();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create an object containing HTML export options.
            HtmlDocumentExporterOptions options = new HtmlDocumentExporterOptions();

            // Set HTML-specific export options.
            options.CssPropertiesExportType = DevExpress.XtraSpreadsheet.Export.Html.CssPropertiesExportType.Style;
            options.Encoding = Encoding.UTF8;

            // Specify the part of the document to be exported to HTML.
            options.SheetIndex = worksheet.Index;
            options.Range = "B11:O28";

            // Export a document to an HTML file with the specified options.
            workbook.ExportToHtml("OutputRange.html", options);

            // Export the entire worksheet to a stream as HTML.
            FileStream htmlStream = new FileStream("OutputWorksheet.html", FileMode.Create);
            workbook.ExportToHtml(htmlStream, worksheet.Index);
            htmlStream.Close();
            #endregion #ExportToHTML

            System.Diagnostics.Process.Start("OutputRange.html");
            System.Diagnostics.Process.Start("OutputWorksheet.html");
        }
    }
}
