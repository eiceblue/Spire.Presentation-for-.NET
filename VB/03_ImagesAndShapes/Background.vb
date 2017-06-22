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
        para_title.Text = "Background"
        para_title.Alignment = TextAlignmentType.Center
        para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
        para_title.TextRanges(0).FontHeight = 36
        para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        para_title.TextRanges(0).Fill.SolidColor.Color = Color.Black
        shape_title.TextFrame.Paragraphs.Append(para_title)

        'add new shape to PPT document
        Dim rec As New RectangleF(presentation.SlideSize.Size.Width / 2 - 300, 155, 600, 300)
        Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)

        shape.ShapeStyle.LineColor.Color = Color.White
        shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

        'add text to shape
        shape.AppendTextFrame("Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine.")

        Dim para As New TextParagraph()
        para.Text = "Spire.Presentation for .NET support PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc."
        shape.TextFrame.Paragraphs.Append(para)

        'set the font and fill style of text
        For Each paragraph As TextParagraph In shape.TextFrame.Paragraphs
            paragraph.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
            paragraph.TextRanges(0).Fill.SolidColor.Color = Color.Black
            paragraph.TextRanges(0).FontHeight = 20
            paragraph.TextRanges(0).LatinFont = New TextFont("Myriad Pro")
            paragraph.Alignment = TextAlignmentType.Left
        Next

        'set spacing after
        shape.TextFrame.Paragraphs(0).SpaceAfter = 80

        'save the document
        presentation.SaveToFile("background.pptx", FileFormat.Pptx2010)
        System.Diagnostics.Process.Start("background.pptx")

    End Sub
End Class