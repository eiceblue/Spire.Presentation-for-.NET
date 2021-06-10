Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace PageSetup
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Set the size of slides
			presentation.SlideSize.Size = New SizeF(600,600)
			presentation.SlideSize.Orientation = SlideOrienation.Portrait
			presentation.SlideSize.Type = SlideSizeType.Custom

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Append new shape
			Dim rec As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 200, 150, 400, 200)
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

			'Add text to shape
			shape.AppendTextFrame("The sample demonstrates how to set slide size.")

			shape.TextFrame.Paragraphs(0).TextRanges(0).LatinFont = New TextFont("Myriad Pro")
			shape.TextFrame.Paragraphs(0).TextRanges(0).FontHeight = 24
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.FillType = FillFormatType.Solid
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.SolidColor.Color = Color.FromArgb(36,64,97)

			'Save the document
			presentation.SaveToFile("PageSetup.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("PageSetup.pptx")
		End Sub
	End Class
End Namespace