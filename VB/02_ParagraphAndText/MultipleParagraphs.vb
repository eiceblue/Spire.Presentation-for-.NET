Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace MultipleParagraphs
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Az.pptx")
			'Access the first slide
			Dim slide As ISlide = presentation.Slides(0)

			' Add an AutoShape of rectangle type
			Dim rec As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 250, 150, 500, 150)
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)

			' Access TextFrame of the AutoShape
			Dim tf As ITextFrameProperties = shape.TextFrame

			' Create Paragraphs and TextRanges with different text formats
			Dim para0 As TextParagraph = tf.Paragraphs(0)
			Dim textRange1 As New TextRange()
			Dim textRange2 As New TextRange()
			para0.TextRanges.Append(textRange1)
			para0.TextRanges.Append(textRange2)

			Dim para1 As New TextParagraph()
			tf.Paragraphs.Append(para1)
			Dim textRange11 As New TextRange()
			Dim textRange12 As New TextRange()
			Dim textRange13 As New TextRange()
			para1.TextRanges.Append(textRange11)
			para1.TextRanges.Append(textRange12)
			para1.TextRanges.Append(textRange13)

			Dim para2 As New TextParagraph()
			tf.Paragraphs.Append(para2)
			Dim textRange21 As New TextRange()
			Dim textRange22 As New TextRange()
			Dim textRange23 As New TextRange()
			para2.TextRanges.Append(textRange21)
			para2.TextRanges.Append(textRange22)
			para2.TextRanges.Append(textRange23)

			For i As Integer = 0 To 2
				For j As Integer = 0 To 2
					tf.Paragraphs(i).TextRanges(j).Text = "TextRange " & j.ToString()
					If j = 0 Then
						tf.Paragraphs(i).TextRanges(j).Fill.FillType = FillFormatType.Solid
						tf.Paragraphs(i).TextRanges(j).Fill.SolidColor.Color = Color.LightBlue
						tf.Paragraphs(i).TextRanges(j).Format.IsBold = TriState.True
						tf.Paragraphs(i).TextRanges(j).FontHeight = 15
					ElseIf j = 1 Then
						tf.Paragraphs(i).TextRanges(j).Fill.FillType = FillFormatType.Solid
						tf.Paragraphs(i).TextRanges(j).Fill.SolidColor.Color = Color.Blue
						tf.Paragraphs(i).TextRanges(j).Format.IsItalic = TriState.True
						tf.Paragraphs(i).TextRanges(j).FontHeight = 18
					End If
				Next j
			Next i


			Dim result As String = "MultipleParagraphs_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)
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