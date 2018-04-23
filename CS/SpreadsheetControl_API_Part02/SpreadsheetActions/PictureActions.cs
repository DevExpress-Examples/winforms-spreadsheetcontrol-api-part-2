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
    public static class PictureActions {
        #region Actions
        public static Action<IWorkbook> InsertPictureAction = InsertPictureValue;
        public static Action<IWorkbook> InsertPictureFromUriAction = InsertPictureFromUriValue;
        public static Action<IWorkbook> MovePictureAction = MovePictureValue;
        public static Action<IWorkbook> RotatePictureAction = RotatePictureValue;
        public static Action<IWorkbook> ChangeZOrderAction = ChangeZOrderValue;
        public static Action<IWorkbook> InsertHyperlinkAction = InsertHyperlinkValue;
        #endregion

        static void InsertPictureValue(IWorkbook workbook)
        {
            #region #insertPicture
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
                worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["D5"]);
                // Insert a picture to fit in the specified range.
                worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", worksheet.Range["B2"]);
                // Insert a picture from the SpreadsheetImageSource at 120 mm from the left, 80 mm from the top, 
                // and resize it to a width of 70 mm and a height of 20 mm, locking the aspect ratio.
                worksheet.Pictures.AddPicture(imageSource, 120, 80, 70, 20, true);
                // Insert the picture to be removed.
                worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", 0, 0);
                // Remove the last inserted picture.
                // Find the Picture by its name. The method returns a collection of Pictures with the same name.
                Picture pic = worksheet.Pictures.GetPicturesByName("Picture 4")[0];
                pic.Delete();
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #insertPicture
        }

        static void InsertPictureFromUriValue(IWorkbook workbook)
        {
            #region #insertpicturefromuri
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
                worksheet.Pictures.AddPicture(imageSource, 100, 40, 200, 180);
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #insertpicturefromuri
            
        }

        static void MovePictureValue(IWorkbook workbook)
        {

            #region #movepicture
            // Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            workbook.Worksheets.ActiveWorksheet.DefaultRowHeight = 20;
            workbook.Worksheets.ActiveWorksheet.DefaultColumnWidth = 20;

            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert pictures.
                Picture pic = worksheet.Pictures.AddPicture("Pictures\\x-spreadsheet.png", worksheet.Cells["A1"]);
                worksheet.Pictures.AddPicture("Pictures\\x-spreadsheet.png", worksheet.Cells["A1"]);
                // Specify picture name.
                pic.Name = "Logo";
                pic.AlternativeText = "Spreadsheet logo";
                // Move a picture.
                pic.Move(30, 50);
                // Move and size the picture with underlying cells. 
                pic.Placement = Placement.MoveAndSize;
                worksheet.Rows[1].Height += 20;
                worksheet.Columns["D"].Width += 20;
                // Move another picture to illustrate OffsetX, OffsetY properties.
                worksheet.Pictures[1].Move(pic.OffsetY, pic.OffsetX);
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #movepicture
        }

        static void RotatePictureValue(IWorkbook workbook)
        {

            #region #rotatepicture
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert picture from the file.
                Picture pic = worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["B5"]);
                // Specify rotation angle.
                pic.Rotation = 30;
                // Increase rotation angle.
                pic.IncrementRotation(15);
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #rotatepicture
        }

        static void ChangeZOrderValue(IWorkbook workbook)
        {

            #region #changezorder
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert pictures.
                Picture pic1 = worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["B2"]);
                Picture pic2 = worksheet.Pictures.AddPicture("Pictures\\x-spreadsheet.png", worksheet.Cells["C5"]);
                // Bring the first picture to front.
                pic1.ZOrderPosition = 3;
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #changezorder
        }

        static void InsertHyperlinkValue(IWorkbook workbook)
        {

            #region #inserthyperlink
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert picture.
                Picture pic = worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["A1"]);
                // Add a hyperlink.
                pic.InsertHyperlink("http://www.devexpress.com/Products/NET/Document-Server/", true);
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #inserthyperlink
        }
    }
}
