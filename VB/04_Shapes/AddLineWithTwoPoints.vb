Imports Spire.Presentation

Namespace AddLineWithTwoPoints
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Get the first slide
			Dim slide As ISlide = ppt.Slides(0)

			'Add line with two points
			Dim line As IAutoShape = slide.Shapes.AppendShape(ShapeType.Line, New PointF(50, 50), New PointF(150, 150))
			line.ShapeStyle.LineColor.Color = Color.Red
			line = slide.Shapes.AppendShape(ShapeType.Line, New PointF(150, 150), New PointF(250, 50))
			line.ShapeStyle.LineColor.Color = Color.Blue

			'Save the document
			Dim result As String = "AddLineWithTwoPoints.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			PresentationDocViewer(result)
		End Sub

	Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
	End Sub
	End Class
End Namespace