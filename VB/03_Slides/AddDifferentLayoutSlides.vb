Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace AddDifferentLayoutSlides
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Remove the default slide
			presentation.Slides.RemoveAt(0)

			'Loop through slide layouts
			For Each type As SlideLayoutType In System.Enum.GetValues(GetType(SlideLayoutType))
				'Append slide by specifing slide layout
				presentation.Slides.Append(type)
			Next type

			'Save the document
			Dim result As String = "AddDifferentLayoutSlides_result.pptx"
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