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

        'set the size of slides
        presentation.SlideSize.Size = New SizeF(500, 600)
        presentation.SlideSize.Orientation = SlideOrienation.Portrait
        presentation.SlideSize.Type = SlideSizeType.[Custom]

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
        para_title.Text = "Page Size"
        para_title.Alignment = TextAlignmentType.Center
        para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
        para_title.TextRanges(0).FontHeight = 36
        para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        para_title.TextRanges(0).Fill.SolidColor.Color = Color.Black
        shape_title.TextFrame.Paragraphs.Append(para_title)

        'append new shape
        Dim rec As New RectangleF(presentation.SlideSize.Size.Width / 2 - 200, 155, 400, 350)
        Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)
        shape.ShapeStyle.LineColor.Color = Color.White
        shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

        'add text to shape
        shape.AppendTextFrame("The sample demonstrates how to set the size of slides.")

        'append new paragraph
        shape.TextFrame.Paragraphs.Append(New TextParagraph())

        'add text to paragraph
        shape.TextFrame.Paragraphs(1).TextRanges.Append(New TextRange("Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. Spire.Presentation for .NET support PPT, PPS, PPTX and PPSX presentation formats."))

        'set the Font
        For Each para As TextParagraph In shape.TextFrame.Paragraphs
            para.TextRanges(0).LatinFont = New TextFont("Myriad Pro")
            para.TextRanges(0).FontHeight = 24
            para.TextRanges(0).Fill.FillType = FillFormatType.Solid
            para.TextRanges(0).Fill.SolidColor.Color = Color.Black
            para.Alignment = TextAlignmentType.Left
        Next

        'save the document
        presentation.SaveToFile("PageProperty.pptx", FileFormat.Pptx2010)
        System.Diagnostics.Process.Start("PageProperty.pptx")

    End Sub
End Class