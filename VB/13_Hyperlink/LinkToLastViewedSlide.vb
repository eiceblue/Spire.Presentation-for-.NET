Imports System.Collections
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace LinkToLastViewedSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()
			'Load the document from disk
			ppt.LoadFromFile("..\..\..\..\..\..\Data\LastViewedSlide.pptx")
			'Get specified slide
			Dim slide As ISlide = ppt.Slides(0)
			'Draw a shape
			Dim autoShape As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(100, 100, 100, 100))
			'Link to last viewed slide show
			autoShape.Click = ClickHyperlink.LastVievedSlide
			'Save the document
			Dim result As String = "GetLastViewedSlide.pptx"
			ppt.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013)
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace