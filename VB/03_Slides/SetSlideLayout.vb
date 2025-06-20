Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetSlideLayout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Remove the first slide
			ppt.Slides.RemoveAt(0)

			'Append a slide and set the layout for slide
			Dim slide As ISlide = ppt.Slides.Append(SlideLayoutType.Title)

			'Add content for Title and Text
			Dim shape As IAutoShape = TryCast(slide.Shapes(0), IAutoShape)
			shape.TextFrame.Text = "Hello Wolrd! ¨C> This is title"

			shape = TryCast(slide.Shapes(1), IAutoShape)
			shape.TextFrame.Text = "E-iceblue Support Team -> This is content"

			'Save the document
			Dim result As String = "SetSlideLayout.pptx"
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