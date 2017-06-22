Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
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
        para_title.Text = "Document Property"
        para_title.Alignment = TextAlignmentType.Center
        para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
        para_title.TextRanges(0).FontHeight = 36
        para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        para_title.TextRanges(0).Fill.SolidColor.Color = Color.Black
        shape_title.TextFrame.Paragraphs.Append(para_title)

        'set the DocumentProperty of PPT document
        presentation.DocumentProperty.Application = "Spire.Presentation"
        presentation.DocumentProperty.Author = "http://www.e-iceblue.com/"
        presentation.DocumentProperty.Company = "E-iceblue"
        presentation.DocumentProperty.Keywords = "Demo File"
        presentation.DocumentProperty.Comments = "This file tests Spire.Presentation."
        presentation.DocumentProperty.Category = "Demo"
        presentation.DocumentProperty.Title = "This is a demo file."
        presentation.DocumentProperty.Subject = "Test"

        'insert image to PPT
        Dim ImageFile2 As String = "..\..\..\..\..\..\Data\Property.png"
        Dim rect1 As New RectangleF(presentation.SlideSize.Size.Width / 2 - 300, 155, 300, 200)
        Dim image As IEmbedImage = presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile2, rect1)
        image.Line.FillType = FillFormatType.None

        'add new shape to PPT document
        Dim rec As New RectangleF(presentation.SlideSize.Size.Width / 2 - 300, 370, 600, 120)
        Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)

        shape.ShapeStyle.LineColor.Color = Color.White
        shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

        'add text to shape
        shape.AppendTextFrame("Spire.Presentation for .NET support PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc.")

        'set the font and fill style of text
        Dim paragraph As TextParagraph = shape.TextFrame.Paragraphs(0)
        paragraph.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        paragraph.TextRanges(0).Fill.SolidColor.Color = Color.Black
        paragraph.TextRanges(0).FontHeight = 20
        paragraph.TextRanges(0).LatinFont = New TextFont("Myriad Pro")
        paragraph.Alignment = TextAlignmentType.Left

        'save the document
        presentation.SaveToFile("DocumentProperty.pptx", FileFormat.Pptx2007)
        System.Diagnostics.Process.Start("DocumentProperty.pptx")

    End Sub
End Class