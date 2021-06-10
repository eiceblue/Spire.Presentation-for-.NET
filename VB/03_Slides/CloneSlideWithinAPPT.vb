Imports Spire.Presentation

Namespace CloneSlideWithinAPPT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\InputTemplate.pptx")

			'Get a list of slides and choose the first slide to be cloned
			Dim slide As ISlide = ppt.Slides(0)

			'Insert the desired slide to the specified index in the same presentation
			Dim index As Integer = 1
			ppt.Slides.Insert(index, slide)

			'Save the document
			Dim result As String = "CloneSlideWithinAPPT.pptx"
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