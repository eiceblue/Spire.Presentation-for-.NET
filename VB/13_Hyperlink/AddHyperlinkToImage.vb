Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace AddHyperlinkToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_5.pptx")

			'Get the first slide.
			Dim slide As ISlide = presentation.Slides(0)

			'Add image to slide.
			Dim rect As New RectangleF(480, 350, 160, 160)
			Dim image As IEmbedImage = slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, "..\..\..\..\..\..\Data\Logo1.png", rect)

			'Add hyperlink to the image.
			Dim hyperlink As New ClickHyperlink("https://www.e-iceblue.com")
			image.Click = hyperlink

			Dim result As String = "Result-AddHyperlinkToImage.pptx"

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