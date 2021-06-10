Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace LinkToASpecificSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Append a slide to it.
			presentation.Slides.Append()

			'Add a shape to the second slide.
			Dim shape As IAutoShape = presentation.Slides(1).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(10, 50, 200, 50))
			shape.Fill.FillType = FillFormatType.None
			shape.Line.FillType = FillFormatType.None
			shape.TextFrame.Text = "Jump to the first slide"

			'Create a hyperlink based on the shape and the text on it, linking to the first slide.
			Dim hyperlink As New ClickHyperlink(presentation.Slides(0))
			shape.Click = hyperlink
			shape.TextFrame.TextRange.ClickAction = hyperlink

			Dim result As String = "Result-LinkToASpecificSlide.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file.
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace