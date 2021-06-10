Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ResetShapeSizeAndPosition
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ShapeTemplate.pptx")

			'Define the original slide size
			Dim currentHeight As Single = ppt.SlideSize.Size.Height
			Dim currentWidth As Single = ppt.SlideSize.Size.Width

			'Change the slide size as A3
			ppt.SlideSize.Type = SlideSizeType.A3

			'Define the new slide size
			Dim newHeight As Single = ppt.SlideSize.Size.Height
			Dim newWidth As Single = ppt.SlideSize.Size.Width

			'Define the ratio from the old and new slide size
			Dim ratioHeight As Single = newHeight / currentHeight
			Dim ratioWidth As Single = newWidth / currentWidth

			'Reset the size and position of the shape on the slide
			For Each slide As ISlide In ppt.Slides
				For Each shape As IShape In slide.Shapes
					shape.Height = shape.Height * ratioHeight
					shape.Width = shape.Width * ratioWidth

					shape.Left = shape.Left * ratioHeight
					shape.Top = shape.Top * ratioWidth
				Next shape
			Next slide

			'Save the document
			Dim result As String = "ResetShapeSizeAndPosition.pptx"
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