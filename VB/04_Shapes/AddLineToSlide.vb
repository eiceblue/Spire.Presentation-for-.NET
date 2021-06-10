Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace AddLineToSlide
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

			'Add a line in the slide
			Dim line As IAutoShape=slide.Shapes.AppendShape(ShapeType.Line, New RectangleF(50, 100, 300, 0))

			'Set color of the line
			line.ShapeStyle.LineColor.Color = Color.Red

			'Save the document
			Dim result As String = "AddLineToSlide_result.pptx"
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