Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace CloneSlideAtTheEnd
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load PPT document from disk
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ChangeSlidePosition.pptx")

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'Append the slide at the end of the document
			presentation.Slides.Append(slide)

			'Save the document
			Dim result As String = "ClonePPTAtTheEnd_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace