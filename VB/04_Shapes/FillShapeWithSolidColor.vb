Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace FillShapeWithSolidColor
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

			'Fill shape with solid color
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.Yellow

			'Set the fill format of line
			shape.Line.FillType = FillFormatType.Solid
			shape.Line.SolidFillColor.Color = Color.Gray

			'Save the document
			Dim result As String = "FillShapeWithSolidColor_result.pptx"
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