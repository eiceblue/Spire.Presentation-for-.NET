Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace ChangeHyperlinkColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document
			Dim presentation As New Presentation()

			'Load file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ChangeHyperlinkColor.pptx")

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'Get the theme of the slide
			Dim theme As Theme = slide.Theme

			'Change the color of hyperlink to red
			theme.ColorScheme.HyperlinkColor.Color = Color.Red

			Dim result As String = "Result-ChangeHyperlinkColor.pptx"

			'Save to file
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace