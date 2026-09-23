Imports System
Imports System.IO
Imports System.Reflection
'#Region "#usings_Sh"
Imports DevExpress.Spreadsheet

'#End Region  ' #usings_Sh
Namespace SpreadsheetControl_API

    Public Module ShapeActions

'#Region "Actions"
        Public InsertShapeAction As Action(Of IWorkbook) = AddressOf InsertShapeValue

        Public InsertShapeFromUriAction As Action(Of IWorkbook) = AddressOf InsertShapeFromUriValue

        Public ModifyShapeAction As Action(Of IWorkbook) = AddressOf ModifyShapeValue

'#End Region
        Private Sub InsertShapeValue(ByVal workbook As IWorkbook)
'#Region "#insertshape"
            Dim imageStream As Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("SpreadsheetControl_API.Pictures.x-spreadsheet.png")
            Dim imageSource As SpreadsheetImageSource = SpreadsheetImageSource.FromStream(imageStream)
            workbook.BeginUpdate()
            ' Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert a picture from a file so that its top left corner is in the specified cell.
                ' By default the picture is named Picture1.. PictureNN.
                worksheet.Shapes.AddPicture("Pictures\x-docserver.png", worksheet.Cells("D5"))
                ' Insert a picture to fit in the specified range.
                worksheet.Shapes.AddPicture("Pictures\x-docserver.png", worksheet.Range("B2"))
                ' Insert a picture from the SpreadsheetImageSource at 120 mm from the left, 80 mm from the top, 
                ' and resize it to a width of 70 mm and a height of 20 mm, locking the aspect ratio.
                worksheet.Shapes.AddPicture(imageSource, 120, 80, 70, 20, True)
                ' Insert the picture to be removed.
                worksheet.Shapes.AddPicture("Pictures\x-docserver.png", 0, 0)
                ' Remove the last inserted picture.
                ' Find the shape by its name. The method returns a collection of shapes with the same name.
                Dim picShape As Shape = worksheet.Shapes.GetShapesByName("Picture4")(0)
                picShape.Delete()
            Finally
                workbook.EndUpdate()
            End Try
'#End Region  ' #insertshape
        End Sub

        Private Sub InsertShapeFromUriValue(ByVal workbook As IWorkbook)
'#Region "#insertshapefromuri"
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
                worksheet.Shapes.AddPicture(imageSource, 100, 40, 200, 180)
            Finally
                workbook.EndUpdate()
            End Try
'#End Region  ' #insertshapefromuri
        End Sub

        Private Sub ModifyShapeValue(ByVal workbook As IWorkbook)
'#Region "#modifyshape"
            ' Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert pictures from the file.
                Dim pic As Picture = worksheet.Shapes.AddPicture("Pictures\x-docserver.png", worksheet.Cells("A1"))
                worksheet.Shapes.AddPicture("Pictures\x-spreadsheet.png", worksheet.Cells("D5"))
                ' Specify picture name and draw a border.
                pic.Name = "Logo"
                pic.AlternativeText = "Document Server logo"
                pic.BorderWidth = 1
                pic.BorderColor = DevExpress.Utils.DXColor.Black
                ' Move a picture.
                pic.Move(20, 30)
                ' Change picture behavior so it will move and size with underlying cells. 
                pic.Placement = Placement.MoveAndSize
                worksheet.Rows(5).Height += 10
                worksheet.Columns("D").Width += 10
                ' Specify rotation angle.
                pic.Rotation = 30
                ' Increase rotation angle.
                pic.IncrementRotation(-30)
                ' Add a hyperlink.
                pic.InsertHyperlink("http://www.devexpress.com/Products/NET/Document-Server/", True)
                worksheet.Shapes(1).InsertHyperlink("http://www.devexpress.com/Products/NET/Controls/WinForms/Spreadsheet/", True)
            Finally
                workbook.EndUpdate()
            End Try
'#End Region  ' #modifyshape
        End Sub
    End Module
End Namespace
