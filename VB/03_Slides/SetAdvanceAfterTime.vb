Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetAdvanceAfterTime
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()

			'Load the document from disk
			ppt.LoadFromFile("..\..\..\..\..\..\..\Data\SetTransitions.pptx")

			'Traverse all slides
			For i As Integer = 0 To ppt.Slides.Count - 1
				ppt.Slides(i).SlideShowTransition.AdvanceOnClick = True

				'Set the time
				ppt.Slides(i).SlideShowTransition.AdvanceAfterTime = 5000
			Next i

			Dim result As String = "Result.pptx"
			'Save the document
			ppt.SaveToFile(result, FileFormat.Pptx2010)

			PresentationDocViewer(result)
		End Sub

		Private Shared Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace