Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Threading.Tasks
#Region "#usings_Sh"
Imports DevExpress.Spreadsheet
#End Region ' #usings_Sh

Namespace SpreadsheetControl_API
    Public NotInheritable Class PictureActions

        Private Sub New()
        End Sub

        #Region "Actions"
        Public Shared InsertPictureAction As Action(Of IWorkbook) = AddressOf InsertPictureValue
        Public Shared InsertPictureFromUriAction As Action(Of IWorkbook) = AddressOf InsertPictureFromUriValue
        Public Shared MovePictureAction As Action(Of IWorkbook) = AddressOf MovePictureValue
        Public Shared RotatePictureAction As Action(Of IWorkbook) = AddressOf RotatePictureValue
        Public Shared ChangeZOrderAction As Action(Of IWorkbook) = AddressOf ChangeZOrderValue
        Public Shared InsertHyperlinkAction As Action(Of IWorkbook) = AddressOf InsertHyperlinkValue
        #End Region

        Private Shared Sub InsertPictureValue(ByVal workbook As IWorkbook)
            '            #Region "#insertPicture"
            Dim imageStream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("x-spreadsheet.png")
            Dim imageSource As SpreadsheetImageSource = SpreadsheetImageSource.FromStream(imageStream)
            workbook.BeginUpdate()
            ' Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert a picture from a file so that its top left corner is in the specified cell.
                ' By default the picture is named Picture1.. PictureNN.
                worksheet.Pictures.AddPicture("Pictures\x-docserver.png", worksheet.Cells("D5"))
                ' Insert a picture to fit in the specified range.
                worksheet.Pictures.AddPicture("Pictures\x-docserver.png", worksheet.Range("B2"))
                ' Insert a picture from the SpreadsheetImageSource at 120 mm from the left, 80 mm from the top, 
                ' and resize it to a width of 70 mm and a height of 20 mm, locking the aspect ratio.
                worksheet.Pictures.AddPicture(imageSource, 120, 80, 70, 20, True)
                ' Insert the picture to be removed.
                worksheet.Pictures.AddPicture("Pictures\x-docserver.png", 0, 0)
                ' Remove the last inserted picture.
                ' Find the Picture by its name. The method returns a collection of Pictures with the same name.
                Dim pic As Picture = worksheet.Pictures.GetPicturesByName("Picture 4")(0)
                pic.Delete()
            Finally
                workbook.EndUpdate()
            End Try
'            #End Region ' #insertPicture
        End Sub

        Private Shared Sub InsertPictureFromUriValue(ByVal workbook As IWorkbook)
'            #Region "#insertpicturefromuri"
            Dim imageUri As String = "http://www.devexpress.com/Products/NET/Document-Server/i/Unit-Conversion.png"
            ' Create an image from Uri.
            Dim imageSource As SpreadsheetImageSource = SpreadsheetImageSource.FromUri(imageUri, workbook)
            ' Set the measurement unit to point.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point

            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert a picture from the SpreadsheetImageSource at 100 pt from the left, 40 pt from the top, 
                ' and resize it to a width of 200 pt and a height of 180 pt.
                worksheet.Pictures.AddPicture(imageSource, 100, 40, 200, 180)
            Finally
                workbook.EndUpdate()
            End Try
'            #End Region ' #insertpicturefromuri

        End Sub

        Private Shared Sub MovePictureValue(ByVal workbook As IWorkbook)

'            #Region "#movepicture"
            ' Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter
            workbook.Worksheets.ActiveWorksheet.DefaultRowHeight = 20
            workbook.Worksheets.ActiveWorksheet.DefaultColumnWidth = 20

            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert pictures.
                Dim pic As Picture = worksheet.Pictures.AddPicture("Pictures\x-spreadsheet.png", worksheet.Cells("A1"))
                worksheet.Pictures.AddPicture("Pictures\x-spreadsheet.png", worksheet.Cells("A1"))
                ' Specify picture name.
                pic.Name = "Logo"
                pic.AlternativeText = "Spreadsheet logo"
                ' Move a picture.
                pic.Move(30, 50)
                ' Move and size the picture with underlying cells. 
                pic.Placement = Placement.MoveAndSize
                worksheet.Rows(1).Height += 20
                worksheet.Columns("D").Width += 20
                ' Move another picture to illustrate OffsetX, OffsetY properties.
                worksheet.Pictures(1).Move(pic.OffsetY, pic.OffsetX)
            Finally
                workbook.EndUpdate()
            End Try
'            #End Region ' #movepicture
        End Sub

        Private Shared Sub RotatePictureValue(ByVal workbook As IWorkbook)

'            #Region "#rotatepicture"
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert picture from the file.
                Dim pic As Picture = worksheet.Pictures.AddPicture("Pictures\x-docserver.png", worksheet.Cells("B5"))
                ' Specify rotation angle.
                pic.Rotation = 30
                ' Increase rotation angle.
                pic.IncrementRotation(15)
            Finally
                workbook.EndUpdate()
            End Try
'            #End Region ' #rotatepicture
        End Sub

        Private Shared Sub ChangeZOrderValue(ByVal workbook As IWorkbook)

'            #Region "#changezorder"
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert pictures.
                Dim pic1 As Picture = worksheet.Pictures.AddPicture("Pictures\x-docserver.png", worksheet.Cells("B2"))
                Dim pic2 As Picture = worksheet.Pictures.AddPicture("Pictures\x-spreadsheet.png", worksheet.Cells("C5"))
                ' Bring the first picture to front.
                pic1.ZOrderPosition = 3
            Finally
                workbook.EndUpdate()
            End Try
'            #End Region ' #changezorder
        End Sub

        Private Shared Sub InsertHyperlinkValue(ByVal workbook As IWorkbook)

'            #Region "#inserthyperlink"
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert picture.
                Dim pic As Picture = worksheet.Pictures.AddPicture("Pictures\x-docserver.png", worksheet.Cells("A1"))
                ' Add a hyperlink.
                pic.InsertHyperlink("http://www.devexpress.com/Products/NET/Document-Server/", True)
            Finally
                workbook.EndUpdate()
            End Try
'            #End Region ' #inserthyperlink
        End Sub
    End Class
End Namespace
