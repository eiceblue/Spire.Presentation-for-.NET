Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing.Animation
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace Hyperlinks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Set background Image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)

			'Add new shape to PPT document
			Dim rec As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 255, 120, 500, 280)
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			shape.Line.Width = 0

			'Add some paragraphs with hyperlinks
			Dim para1 As New TextParagraph()
			Dim tr As New TextRange("E-iceblue")
			tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			tr.Fill.SolidColor.Color = Color.Blue
			para1.TextRanges.Append(tr)
			para1.Alignment = TextAlignmentType.Center
			shape.TextFrame.Paragraphs.Append(para1)
			shape.TextFrame.Paragraphs.Append(New TextParagraph())

			'Add some paragraphs with hyperlinks
			Dim para2 As New TextParagraph()
			Dim tr1 As New TextRange("Click to know more about Spire.Presentation.")
			tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html"
			para2.TextRanges.Append(tr1)
			shape.TextFrame.Paragraphs.Append(para2)
			shape.TextFrame.Paragraphs.Append(New TextParagraph())

			Dim para3 As New TextParagraph()
			Dim tr2 As New TextRange("Click to visit E-iceblue Home page.")
			tr2.ClickAction.Address = "https://www.e-iceblue.com/"
			para3.TextRanges.Append(tr2)
			shape.TextFrame.Paragraphs.Append(para3)
			shape.TextFrame.Paragraphs.Append(New TextParagraph())

			Dim para4 As New TextParagraph()
			Dim tr3 As New TextRange("Click to go to the forum to raise questions.")
			tr3.ClickAction.Address = "https://www.e-iceblue.com/forum/components-f5.html"
			para4.TextRanges.Append(tr3)
			shape.TextFrame.Paragraphs.Append(para4)
			shape.TextFrame.Paragraphs.Append(New TextParagraph())

			Dim para5 As New TextParagraph()
			Dim tr4 As New TextRange("Click to contact our sales team via email.")
			tr4.ClickAction.Address = "mailto:sales@e-iceblue.com"
			para5.TextRanges.Append(tr4)
			shape.TextFrame.Paragraphs.Append(para5)
			shape.TextFrame.Paragraphs.Append(New TextParagraph())

			Dim para6 As New TextParagraph()
			Dim tr5 As New TextRange("Click to contact our support team via email.")
			tr5.ClickAction.Address = "mailto:support@e-iceblue.com"
			para6.TextRanges.Append(tr5)
			shape.TextFrame.Paragraphs.Append(para6)

			For Each para As TextParagraph In shape.TextFrame.Paragraphs
				If Not String.IsNullOrEmpty(para.Text) Then
					para.TextRanges(0).LatinFont = New TextFont("Lucida Sans Unicode")
					para.TextRanges(0).FontHeight = 20
				End If

			Next para

			'Save the document
			presentation.SaveToFile("hyperlink_result.pptx", FileFormat.Pptx2010)
			Process.Start("hyperlink_result.pptx")
		End Sub
	End Class
End Namespace