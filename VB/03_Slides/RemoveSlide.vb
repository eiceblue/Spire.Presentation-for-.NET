Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace RemoveSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\RemoveSlide.pptx")

			'Remove slide by index
			presentation.Slides.RemoveAt(0)

			'Remove slide by its reference
			Dim slide As ISlide = presentation.Slides(1)
			presentation.Slides.Remove(slide)

			'Save the document
			Dim result As String = "RemoveSlide_result.pptx"
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
