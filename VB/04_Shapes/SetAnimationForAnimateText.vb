Imports Spire.Presentation

Namespace SetAnimationForAnimateText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\..\Data\Animation.pptx")

			'Set the AnimateType as Letter
			ppt.Slides(0).Timeline.MainSequence(0).IterateType = Spire.Presentation.Drawing.TimeLine.AnimateType.Letter

			'Set the IterateTimeValue for the animate text
			ppt.Slides(0).Timeline.MainSequence(0).IterateTimeValue = 10

			'Save the document
			Dim result As String = "SetAnimationForAnimateText.pptx"
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