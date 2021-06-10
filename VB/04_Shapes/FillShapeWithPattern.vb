Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace FillShapeWithPattern
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'Add a rectangle
			Dim rect As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 50, 100, 100, 100)
			Dim shape As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect)

			'Set the pattern fill format 
			shape.Fill.FillType = FillFormatType.Pattern
			shape.Fill.Pattern.PatternType = PatternFillType.Trellis
			shape.Fill.Pattern.BackgroundColor.Color = Color.DarkGray
			shape.Fill.Pattern.ForegroundColor.Color = Color.Yellow

			'Set the fill format of line
			shape.Line.FillType = FillFormatType.Solid
			shape.Line.SolidFillColor.Color = Color.Transparent

			'Save the document
			Dim result As String = "FillShapeWithPattern_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace