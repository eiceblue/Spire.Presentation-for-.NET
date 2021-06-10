Imports Spire.Presentation

Namespace ReorderOverlappingShapes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\OverlappingShapes.pptx")

			'Get the first shape of the first slide
			Dim shape As IShape = ppt.Slides(0).Shapes(0)
			'Change the shape's zorder
			ppt.Slides(0).Shapes.ZOrder(1, shape)

			'Save the document
			Dim result As String = "ReorderOverlappingShapes.pptx"
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