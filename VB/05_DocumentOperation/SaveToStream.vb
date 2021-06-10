Imports System.IO
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SaveToStream
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PowerPoint file and save it to stream
			Dim presentation As New Presentation()

			'Set background Image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Append new shape
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 100, 600, 150))
			shape.Fill.FillType = FillFormatType.None
			shape.ShapeStyle.LineColor.Color = Color.White

			'Add text to shape
			shape.TextFrame.Text = "This demo shows how to Create PowerPoint file and save it to Stream."
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.FillType = FillFormatType.Solid
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.SolidColor.Color = Color.Black
			shape.TextFrame.Paragraphs(0).TextRanges(0).FontHeight = 30

			'Save to Stream
			Dim to_stream As New FileStream("SaveToStream.pptx", FileMode.Create)
			presentation.SaveToFile(to_stream, FileFormat.Pptx2013)
			to_stream.Close()
			PresentationDocViewer("SaveToStream.pptx")
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace