Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ClonePPTAtEndOfAnother
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load source document from disk
			Dim sourcePPT As New Presentation()
			sourcePPT.LoadFromFile("..\..\..\..\..\..\Data\ChangeSlidePosition.pptx")

			'Load destination document from disk
			Dim destPPT As New Presentation()
			destPPT.LoadFromFile("..\..\..\..\..\..\Data\PPTSample_N.pptx")

			'Loop through all slides of source document
			For Each slide As ISlide In sourcePPT.Slides
				'Append the slide at the end of destination document
				destPPT.Slides.Append(slide)
			Next slide

			'Save the document
			Dim result As String = "ClonePPTAtEndOfAnother_result.pptx"
			destPPT.SaveToFile(result, FileFormat.Pptx2013)

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