Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Collections
Imports Spire.Presentation.Drawing.Animation

Namespace SlideTitle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\InputTemplate.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			'Get the title of the first slide
			Dim slideTitle As String = slide.Title

			'Set the title of the second slide
			presentation.Slides(1).Title = "Second Slide"

			Dim result As String = "SlideTitle_result.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file.
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