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

        'add new shape to PPT document
        Dim rec As New RectangleF(presentation.SlideSize.Size.Width / 2 - 250, 80, 500, 150)
        Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)

        shape.ShapeStyle.LineColor.Color = Color.White
        shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

        'add text to shape
        shape.AppendTextFrame("Hello World!")

        'set the font and fill style of text
        Dim textRange As TextRange = shape.TextFrame.TextRange
        textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
        textRange.Fill.SolidColor.Color = Color.Black
        textRange.FontHeight = 72
        textRange.LatinFont = New TextFont("Myriad Pro Light")

        'save the document
        presentation.SaveToFile("hello.pptx", FileFormat.Pptx2010)
        System.Diagnostics.Process.Start("hello.pptx")

    End Sub
End Class