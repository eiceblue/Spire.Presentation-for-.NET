Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace RemoveNoteAtSpecificSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\RemoveNoteFromSlides.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'Get note slide
			Dim note As NotesSlide = slide.NotesSlide
			'Clear note text
			note.NotesTextFrame.Text = ""

			Dim result As String = "RemoveNotesAtSpecificSlide_result.pptx"
			'Save the PPT to PDF file format
			presentation.SaveToFile(result, FileFormat.Pptx2007)

			Viewer(result)
		End Sub

		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace