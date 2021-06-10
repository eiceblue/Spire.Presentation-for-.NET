Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace HyperlinkOutlineStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Add new shape to PPT document
			Dim rec As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 255, 120, 400, 100)
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)
			shape.Fill.FillType =FillFormatType.None
			shape.Line.FillType = FillFormatType.None

			'Add a paragraph with hyperlink
			Dim para1 As New TextParagraph()
			Dim tr1 As New TextRange("Click to know more about Spire.Presentation")
			tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html"
			para1.TextRanges.Append(tr1)

			'Set the format of textrange
			tr1.Format.FontHeight = 20f
			tr1.IsItalic = TriState.True

			'Set the outline format of textrange
			tr1.TextLineFormat.FillFormat.FillType = FillFormatType.Solid
			tr1.TextLineFormat.FillFormat.SolidFillColor.Color = Color.LightSeaGreen
			tr1.TextLineFormat.JoinStyle = LineJoinType.Round
			tr1.TextLineFormat.Width = 2f

			'Add the paragraph to shape
			shape.TextFrame.Paragraphs.Append(para1)
			shape.TextFrame.Paragraphs.Append(New TextParagraph())

			'Save the document
			Dim result As String = "HyperlinkOutlineStyle.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace