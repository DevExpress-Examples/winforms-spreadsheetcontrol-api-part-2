using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
#region #usings_Sh
using DevExpress.Spreadsheet;
#endregion #usings_Sh

namespace SpreadsheetControl_API
{
    public static class ShapeActions {
        #region Actions
        public static Action<IWorkbook> InsertShapeAction = InsertShapeValue;
        public static Action<IWorkbook> InsertShapeFromUriAction = InsertShapeFromUriValue;
        public static Action<IWorkbook> ModifyShapeAction = ModifyShapeValue;
        #endregion

        static void InsertShapeValue(IWorkbook workbook)
        {
            #region #insertshape
            Stream imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("SpreadsheetControl_API.Pictures.x-spreadsheet.png");
            SpreadsheetImageSource imageSource = SpreadsheetImageSource.FromStream(imageStream);
            workbook.BeginUpdate();
            // Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert a picture from a file so that its top left corner is in the specified cell.
                // By default the picture is named Picture1.. PictureNN.
                worksheet.Shapes.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["D5"]);
                // Insert a picture to fit in the specified range.
                worksheet.Shapes.AddPicture("Pictures\\x-docserver.png", worksheet.Range["B2"]);
                // Insert a picture from the SpreadsheetImageSource at 120 mm from the left, 80 mm from the top, 
                // and resize it to a width of 70 mm and a height of 20 mm, locking the aspect ratio.
                worksheet.Shapes.AddPicture(imageSource, 120, 80, 70, 20, true);
                // Insert the picture to be removed.
                worksheet.Shapes.AddPicture("Pictures\\x-docserver.png", 0,0);
                // Remove the last inserted picture.
                // Find the shape by its name. The method returns a collection of shapes with the same name.
                Shape picShape = worksheet.Shapes.GetShapesByName("Picture4")[0];
                picShape.Delete();
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #insertshape
        }

        static void InsertShapeFromUriValue(IWorkbook workbook)
        {
            #region #insertshapefromuri
            string imageUri = "http://www.devexpress.com/Products/NET/Document-Server/i/Unit-Conversion.png";
            // Create an image from Uri.
            SpreadsheetImageSource imageSource = SpreadsheetImageSource.FromUri(imageUri, workbook);
            // Set the measurement unit to point.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;
                        
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert a picture from the SpreadsheetImageSource at 100 pt from the left, 40 pt from the top, 
                // and resize it to a width of 200 pt and a height of 180 pt.
                worksheet.Shapes.AddPicture(imageSource, 100, 40, 200, 180 );
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #insertshapefromuri
            
        }

        static void ModifyShapeValue(IWorkbook workbook)
        {

            #region #modifyshape
            // Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert pictures from the file.
                Picture pic = worksheet.Shapes.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["A1"]);
                worksheet.Shapes.AddPicture("Pictures\\x-spreadsheet.png", worksheet.Cells["D5"]);
                // Specify picture name and draw a border.
                pic.Name = "Logo";
                pic.AlternativeText = "Document Server logo";
                pic.BorderWidth = 1;
                pic.BorderColor = DevExpress.Utils.DXColor.Black;
                // Move a picture.
                pic.Move(20, 30);
                // Change picture behavior so it will move and size with underlying cells. 
                pic.Placement = Placement.MoveAndSize;
                worksheet.Rows[5].Height += 10;
                worksheet.Columns["D"].Width += 10;
                // Specify rotation angle.
                pic.Rotation = 30;
                // Increase rotation angle.
                pic.IncrementRotation(-30);
                // Add a hyperlink.
                pic.InsertHyperlink("http://www.devexpress.com/Products/NET/Document-Server/", true);
                worksheet.Shapes[1].InsertHyperlink("http://www.devexpress.com/Products/NET/Controls/WinForms/Spreadsheet/", true);
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #modifyshape
        }



    }
}
