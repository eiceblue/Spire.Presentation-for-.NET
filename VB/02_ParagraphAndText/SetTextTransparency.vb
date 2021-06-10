Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetTextTransparency
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

			'Add a shape
			Dim textboxShape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(100, 100, 300, 120))
			textboxShape.ShapeStyle.LineColor.Color = Color.Transparent
			textboxShape.Fill.FillType = FillFormatType.None

			'Remove default blank paragraphs
			textboxShape.TextFrame.Paragraphs.Clear()

			'Add three paragraphs, apply color with different alpha values to text
			Dim alpha As Integer = 55
			For i As Integer = 0 To 2
				textboxShape.TextFrame.Paragraphs.Append(New TextParagraph())
				textboxShape.TextFrame.Paragraphs(i).TextRanges.Append(New TextRange("Text Transparency"))
				textboxShape.TextFrame.Paragraphs(i).TextRanges(0).Fill.FillType = FillFormatType.Solid
				textboxShape.TextFrame.Paragraphs(i).TextRanges(0).Fill.SolidColor.Color = Color.FromArgb(alpha, Color.Purple)
				alpha += 100
			Next i

			'Save the document
			Dim result As String = "SetTextTransparency.pptx"
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