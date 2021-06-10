Imports Spire.Presentation

Namespace LoopPresentPPT
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

			'Set the Boolean value of ShowLoop as true
			ppt.ShowLoop = True

			'Set the PowerPoint document to show animation and narration
			ppt.ShowAnimation = True
			ppt.ShowNarration = True
			'Use slide transition timings to advance slide
			ppt.UseTimings = True

			'Save the document
			Dim result As String = "LoopPresentPPT.pptx"
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