Imports Spire.Presentation

Namespace SplitPPT
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

			For i As Integer = 0 To ppt.Slides.Count - 1
				'Initialize another instance of Presentation, and remove the blank slide
				Dim newppt As New Presentation()
				newppt.Slides.RemoveAt(0)

				'Append the specified slide from old presentation to the new one
				newppt.Slides.Append(ppt.Slides(i))

				'Save the document
				Dim result As String = String.Format("SplitPPT-{0}.pptx", i)
				newppt.SaveToFile(result, FileFormat.Pptx2010)
				PresentationDocViewer(result)
			Next i
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace