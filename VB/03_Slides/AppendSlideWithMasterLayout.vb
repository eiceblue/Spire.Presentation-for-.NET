Imports Spire.Presentation
Imports Spire.Presentation.Collections
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace AppendSlideWithMasterLayout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AppendSlideWithMasterLayout.pptx")

			'Get the master
			Dim master As IMasterSlide = presentation.Masters(0)

			'Get master layout slides
			Dim masterLayouts As IMasterLayouts = master.Layouts
			Dim layoutSlide As ActiveSlide = TryCast(masterLayouts(1), ActiveSlide)

			'Append a rectangle to the layout slide
			Dim shape As IAutoShape = layoutSlide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(10, 50, 100, 80))

			'Add a text into the shape and set the style
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			shape.AppendTextFrame("Layout slide 1")
			shape.TextFrame.Paragraphs(0).TextRanges(0).LatinFont = New TextFont("Arial Black")
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.FillType = FillFormatType.Solid
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.SolidColor.Color = Color.CadetBlue

			'Append new slide with master layout
			presentation.Slides.Append(presentation.Slides(0), master.Layouts(1))

			'Another way to append new slide with master layout
			presentation.Slides.Insert(2, presentation.Slides(1), master.Layouts(1))

			'Save the document
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("output.pptx")
		End Sub
	End Class
End Namespace
