Imports Spire.Presentation

Namespace MergeSelectedSlides
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

			'Load two PPT files
			Dim ppt1 As New Presentation("..\..\..\..\..\..\Data\InputTemplate.pptx", FileFormat.Pptx2013)
			Dim ppt2 As New Presentation("..\..\..\..\..\..\Data\TextTemplate.pptx", FileFormat.Pptx2013)

			'Append all slides in ppt1 to ppt
			For i As Integer = 0 To ppt1.Slides.Count - 1
				ppt.Slides.Append(ppt1.Slides(i))
			Next i

			'Append the second slide in ppt2 to ppt
			ppt.Slides.Append(ppt2.Slides(1))

			'Save the document
			Dim result As String = "MergeSelectedSlides.pptx"
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