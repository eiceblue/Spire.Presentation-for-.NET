Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports Spire.Presentation.Drawing

Public Class Form1

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click

        'create PPT document
        Dim presentation As New Presentation()

        'set background Image
        Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
        Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
        presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
        presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

        'add title
        Dim rec_title As New RectangleF(presentation.SlideSize.Size.Width / 2 - 200, 70, 400, 50)
        Dim shape_title As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec_title)
        shape_title.ShapeStyle.LineColor.Color = Color.White
        shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
        Dim para_title As New TextParagraph()
        para_title.Text = "Video"
        para_title.Alignment = TextAlignmentType.Center
        para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
        para_title.TextRanges(0).FontHeight = 36
        para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        para_title.TextRanges(0).Fill.SolidColor.Color = Color.Black
        shape_title.TextFrame.Paragraphs.Append(para_title)

        'insert video into the document
        Dim videoRect As New RectangleF(100, 130, 100, 100)
        Dim video As IVideo = presentation.Slides(0).Shapes.AppendVideoMedia(Path.GetFullPath("..\..\..\..\..\..\Data\Spire.Doc Word to HTML.mp4"), videoRect)
        video.PictureFill.Picture.Url = "..\..\..\..\..\..\Data\video.bmp"

        'add new shape to PPT document
        Dim rec As New RectangleF(presentation.SlideSize.Size.Width / 2 - 300, 255, 600, 150)
        Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)

        shape.ShapeStyle.LineColor.Color = Color.White
        shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

        'add text to shape
        shape.AppendTextFrame("Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine.")

        'set the font
        Dim paragraph As TextParagraph = shape.TextFrame.Paragraphs(0)
        paragraph.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        paragraph.TextRanges(0).Fill.SolidColor.Color = Color.Black
        paragraph.TextRanges(0).FontHeight = 20
        paragraph.TextRanges(0).LatinFont = New TextFont("Myriad Pro")
        paragraph.Alignment = TextAlignmentType.Left

        'save the document
        presentation.SaveToFile("video.pptx", FileFormat.Pptx2010)
        System.Diagnostics.Process.Start("video.pptx")

    End Sub
End Class