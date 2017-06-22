Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

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
        para_title.Text = "Text Color"
        para_title.Alignment = TextAlignmentType.Center
        para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
        para_title.TextRanges(0).FontHeight = 36
        para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        para_title.TextRanges(0).Fill.SolidColor.Color = Color.Black
        shape_title.TextFrame.Paragraphs.Append(para_title)

        'append new shape
        Dim rec As New RectangleF(presentation.SlideSize.Size.Width / 2 - 300, 155, 600, 300)
        Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)

        'set the LineColor
        shape.ShapeStyle.LineColor.Color = Color.Black

        'set the color and fill style
        shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        shape.Fill.SolidColor.Color = Color.LightGray

        'add text to shape
        shape.AppendTextFrame("Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine.")

        'set the color of text
        Dim textRange As TextRange = shape.TextFrame.TextRange
        textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        textRange.Fill.SolidColor.Color = Color.Green

        textRange.Paragraph.Alignment = TextAlignmentType.Left

        'set the Font of text
        textRange.FontHeight = 21
        textRange.IsItalic = TriState.[True]
        textRange.LatinFont = New TextFont("Myriad Pro")
        textRange.FontHeight = 20

        'set spacing after
        shape.TextFrame.Paragraphs(0).SpaceAfter = 80

        'add another paragraph
        Dim para As New TextParagraph()
        para.Text = "Spire.Presentation for .NET support PPT, PPS, PPTX and PPTX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc."
        shape.TextFrame.Paragraphs.Append(para)
        para.TextRanges(0).LatinFont = New TextFont("Myriad Pro")
        para.TextRanges(0).FontHeight = 20
        para.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        para.TextRanges(0).TextUnderlineType = TextUnderlineType.[Single]

        'save the document
        presentation.SaveToFile("text.pptx", FileFormat.Pptx2007)
        System.Diagnostics.Process.Start("text.pptx")

    End Sub
End Class