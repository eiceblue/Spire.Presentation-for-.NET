Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.IO

Namespace AddAndGetSpeakerNotes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_1.pptx")

			'Get the first slide and in the PowerPoint document.
			Dim slide As ISlide = presentation.Slides(0)

			'Get the NotesSlide in the first slide,if there is no notes, we need to add it firstly.
			Dim ns As NotesSlide = slide.NotesSlide
			If ns Is Nothing Then
				ns = slide.AddNotesSlide()
			End If

			'Add the text string as the notes.
			ns.NotesTextFrame.Text = "Speak notes added by Spire.Presentation"

			Dim content As New StringBuilder()
			content.AppendLine("The speaker notes added by Spire.Presentation is: " & ns.NotesTextFrame.Text)

			Dim result As String = "Result-AddAndGetSpeakerNotes.pptx"
			Dim result1 As String = "Result-AddAndGetSpeakerNotes.txt"

			'Save to PowerPoint file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Get the speaker notes and save to txt file.
			File.WriteAllText(result1,content.ToString())

			'Launch the PowerPoint file.
			PptDocumentViewer(result)

			'Launch the txt file.
			PptDocumentViewer(result1)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace