Imports System.Windows.Forms
Imports Spire.Presentation.Drawing.Animation
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
        para_title.Text = "Bullets"
        para_title.Alignment = TextAlignmentType.Center
        para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
        para_title.TextRanges(0).FontHeight = 36
        para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        para_title.TextRanges(0).Fill.SolidColor.Color = Color.Black
        shape_title.TextFrame.Paragraphs.Append(para_title)

        'append new shape
        Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(110, 155, 400, 280))
        shape.Fill.FillType = FillFormatType.None
        shape.ShapeStyle.LineColor.Color = Color.White
        shape.TextFrame.Paragraphs.RemoveAt(0)

        Dim str As String() = New String() {"Spire.Office for .NET", "Spire.Doc for .NET", "Spire.XLS for .NET", "Spire.PDF for .NET", "Spire.Presentation for .NET", "Spire.Barcode for .NET", _
         "Spire.DataExport for .NET", "Spire.DocViewer for .NET", "Spire.PDFViewer for .NET"}
        For Each txt As String In str
            Dim textParagraph As New TextParagraph()
            textParagraph.Text = txt
            textParagraph.Alignment = TextAlignmentType.Left
            textParagraph.Indent = 35

            'set the Bullets
            textParagraph.BulletType = TextBulletType.Numbered
            textParagraph.BulletStyle = NumberedBulletStyle.BulletRomanLCPeriod
            shape.TextFrame.Paragraphs.Append(textParagraph)
        Next

        'set the font and fill style
        For Each paragraph As TextParagraph In shape.TextFrame.Paragraphs
            paragraph.TextRanges(0).LatinFont = New TextFont("Myriad Pro")
            paragraph.TextRanges(0).FontHeight = 24
            paragraph.TextRanges(0).Fill.FillType = FillFormatType.Solid
            paragraph.TextRanges(0).Fill.SolidColor.Color = Color.Black
        Next

        'save the document
        presentation.SaveToFile("bullets.pptx", FileFormat.Pptx2010)
        System.Diagnostics.Process.Start("bullets.pptx")

    End Sub
End Class