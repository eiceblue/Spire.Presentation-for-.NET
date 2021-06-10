Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ArrangeShapes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ArrangeShape.pptx")

			'Get the specified shape
			Dim shape As IShape = ppt.Slides(0).Shapes(0)

			'Bring the shape forward through SetShapeArrange method
			shape.SetShapeArrange(ShapeArrange.BringForward)

			'Save the document
			Dim result As String = "ArrangeShapes.pptx"
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