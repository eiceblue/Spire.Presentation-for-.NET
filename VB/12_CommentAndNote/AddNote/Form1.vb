Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace AddNote
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\AddNote.pptx")

			Dim slide As ISlide = ppt.Slides(0)

			'Add note slide
			Dim notesSlide As NotesSlide = slide.AddNotesSlide()

			'Add paragraph in the notesSlide
			Dim paragraph As New TextParagraph()
			paragraph.Text = "Tips for making effective presentations:"
			notesSlide.NotesTextFrame.Paragraphs.Append(paragraph)

			paragraph = New TextParagraph()
			paragraph.Text = "Use the slide master feature to create a consistent and simple design template."
			notesSlide.NotesTextFrame.Paragraphs.Append(paragraph)
			'Set the bullet type for the paragraph in notesSlide
			notesSlide.NotesTextFrame.Paragraphs(1).BulletType = TextBulletType.Numbered
			notesSlide.NotesTextFrame.Paragraphs(1).BulletStyle = NumberedBulletStyle.BulletArabicPeriod

			paragraph = New TextParagraph()
			paragraph.Text = "Simplify and limit the number of words on each screen."
			notesSlide.NotesTextFrame.Paragraphs.Append(paragraph)
			notesSlide.NotesTextFrame.Paragraphs(2).BulletType = TextBulletType.Numbered
			notesSlide.NotesTextFrame.Paragraphs(2).BulletStyle = NumberedBulletStyle.BulletArabicPeriod

			paragraph = New TextParagraph()
			paragraph.Text = "Use contrasting colors for text and background."
			notesSlide.NotesTextFrame.Paragraphs.Append(paragraph)
			notesSlide.NotesTextFrame.Paragraphs(3).BulletType = TextBulletType.Numbered
			notesSlide.NotesTextFrame.Paragraphs(3).BulletStyle = NumberedBulletStyle.BulletArabicPeriod

			'Save the file
			ppt.SaveToFile("AddNote.pptx", FileFormat.Pptx2010)
			Process.Start("AddNote.pptx")
		End Sub


	End Class
End Namespace
