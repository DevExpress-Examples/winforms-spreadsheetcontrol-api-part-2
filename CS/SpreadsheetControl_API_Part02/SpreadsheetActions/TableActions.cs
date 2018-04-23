using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetControl_API
{
    public static class TableActions
    {
        #region Actions
        public static Action<IWorkbook> CreateTableAction = CreateTable;
        public static Action<IWorkbook> CustomTableStyleAction = CustomTableStyle;
        #endregion


        static void CreateTable(IWorkbook workbook)
        {
            workbook.BeginUpdate();

            Worksheet worksheet = workbook.Worksheets[0];
            GenerateTableData(worksheet);
            #region #CreateTable
            // Insert a table in the worksheet.
            Table table = worksheet.Tables.Add(worksheet["B2:F5"], true);

            // Format the table by applying a built-in table style.
            table.Style = workbook.TableStyles[BuiltInTableStyleId.TableStyleMedium27];

            // Access table columns and name them.
            TableColumn productColumn = table.Columns[0];
            productColumn.Name = "Product";
            TableColumn priceColumn = table.Columns[1];
            priceColumn.Name = "Price";
            TableColumn quantityColumn = table.Columns[2];
            quantityColumn.Name = "Quantity";
            TableColumn discountColumn = table.Columns[3];
            discountColumn.Name = "Discount";
            TableColumn amountColumn = table.Columns[4]; 
            amountColumn.Name = "Amount";

            // Set the formula to calculate the amount per product 
            // and display results in the "Amount" column.
            amountColumn.Formula = "=[Price]*[Quantity]*(1-[Discount])";

            // Display the total row in the table.
            table.ShowTotals = true;

            // Set the label and function to display the sum of the "Amount" column.
            discountColumn.TotalRowLabel = "Total:";
            amountColumn.TotalRowFunction = TotalRowFunction.Sum;

            // Specify the number format for each column.
            priceColumn.DataRange.NumberFormat = "$#,##0.00";
            discountColumn.DataRange.NumberFormat = "0.0%";
            amountColumn.Range.NumberFormat = "$#,##0.00;$#,##0.00;\"\";@";

            // Specify horizontal alignment for header and total rows of the table.
            table.HeaderRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            table.TotalRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

            // Specify horizontal alignment to display data in all columns except the first one.
            for (int i = 1; i < table.Columns.Count; i++)
            {
                table.Columns[i].DataRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            }

            // Set the width of table columns.
            table.Range.ColumnWidthInCharacters = 10;
            #endregion #CreateTable
            workbook.EndUpdate();
        }

        static void CustomTableStyle(IWorkbook workbook)
        {
            CreateTable(workbook);
            workbook.BeginUpdate();
            Worksheet worksheet = workbook.Worksheets[0];
            #region #CustomTableStyle
            // Access a table.
            Table table = worksheet.Tables[0];

            String styleName = "testTableStyle";

            // If the style under the specified name already exists in the collection,
            if (workbook.TableStyles.Contains(styleName))
            {
                // apply this style to the table.
                table.Style = workbook.TableStyles[styleName];
            }
            else
            {
                // Add a new table style under the "testTableStyle" name to the TableStyles collection.
                TableStyle customTableStyle = workbook.TableStyles.Add("testTableStyle");

                // Modify the required formatting characteristics of the table style. 
                // Specify the format for different table elements.
                customTableStyle.BeginUpdate();
                try
                {
                    customTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Font.Color = Color.FromArgb(107, 107, 107);

                    // Specify formatting characteristics for the table header row. 
                    TableStyleElement headerRowStyle = customTableStyle.TableStyleElements[TableStyleElementType.HeaderRow];
                    headerRowStyle.Fill.BackgroundColor = Color.FromArgb(64, 66, 166);
                    headerRowStyle.Font.Color = Color.White;
                    headerRowStyle.Font.Bold = true;

                    // Specify formatting characteristics for the table total row. 
                    TableStyleElement totalRowStyle = customTableStyle.TableStyleElements[TableStyleElementType.TotalRow];
                    totalRowStyle.Fill.BackgroundColor = Color.FromArgb(115, 193, 211);
                    totalRowStyle.Font.Color = Color.White;
                    totalRowStyle.Font.Bold = true;

                    // Specify banded row formatting for the table.
                    TableStyleElement secondRowStripeStyle = customTableStyle.TableStyleElements[TableStyleElementType.SecondRowStripe];
                    secondRowStripeStyle.Fill.BackgroundColor = Color.FromArgb(234, 234, 234);
                    secondRowStripeStyle.StripeSize = 1;
                }
                finally
                {
                    customTableStyle.EndUpdate();
                }
                // Apply the created custom style to the table.
                table.Style = customTableStyle;
            }
            #endregion #CustomTableStyle
            workbook.EndUpdate();

        }

        public static void GenerateTableData(Worksheet sheet)
        {
            sheet.Cells["B3"].SetValue("Chocolade");
            sheet.Cells["B4"].SetValue("Konbu");
            sheet.Cells["B5"].SetValue("Geitost");
            sheet.Cells["C3"].SetValue(5.0);
            sheet.Cells["C4"].SetValue(9.0);
            sheet.Cells["C5"].SetValue(15.0);
            sheet.Cells["D3"].SetValue(15);
            sheet.Cells["D4"].SetValue(55);
            sheet.Cells["D5"].SetValue(70);
            sheet.Cells["E3"].SetValue(0.03);
            sheet.Cells["E4"].SetValue(0.1);
            sheet.Cells["E5"].SetValue(0.07);
        }
    }
}
