Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetTextMargins
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			ppt.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			ppt.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Append a new shape
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 100, 450, 150))

			'Set margins for text inside shapes
			shape.Fill.FillType = FillFormatType.None
			shape.ShapeStyle.LineColor.Color = Color.LightBlue
			shape.TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Justify
			shape.TextFrame.Text = "Using Spire.Presentation, developers will find an easy and effective method to create, read, write, modify, convert and print PowerPoint files on any .Net platform. It's worthwhile for you to try this amazing product."
			shape.TextFrame.Paragraphs(0).TextRanges(0).LatinFont = New TextFont("Arial Rounded MT Bold")
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.FillType = FillFormatType.Solid
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.SolidColor.Color = Color.Black

			'Set the margins for the text frame
			shape.TextFrame.MarginTop = 10
			shape.TextFrame.MarginBottom = 35
			shape.TextFrame.MarginLeft = 15
			shape.TextFrame.MarginRight = 30

			'Save the document
			Dim result As String = "SetTextMargins.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
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