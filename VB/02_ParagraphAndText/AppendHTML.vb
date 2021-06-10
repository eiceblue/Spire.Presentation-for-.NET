Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace AppendHTML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\AppendHTML.pptx")
			'Add a shape 
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(150, 100, 200, 200))

			'Clear default paragraphs 
			shape.TextFrame.Paragraphs.Clear()

			Dim code As String = "<html><body><p>This is a paragraph</p></body></html>"

			'Append HTML, and generate a paragraph with default style in PPT document.
			shape.TextFrame.Paragraphs.AddFromHtml(code)
			Dim codeColor As String = "<html><body><p style="" color:black "">This is a paragraph</p></body></html>"
			'Append HTML with black setting
			shape.TextFrame.Paragraphs.AddFromHtml(codeColor)

			'Add another shape
			Dim shape1 As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(350, 100, 200, 200))

			'Clear default paragraph 
			shape1.TextFrame.Paragraphs.Clear()

			'Change the fill format of shape
			shape1.Fill.FillType = FillFormatType.Solid
			shape1.Fill.SolidColor.Color = Color.White

			'Append HTML
			shape1.TextFrame.Paragraphs.AddFromHtml(code)
			Dim par As TextParagraph = shape1.TextFrame.Paragraphs(0)
			'Change the fill color for paragraph
			For Each tr As TextRange In par.TextRanges
				tr.Fill.FillType = FillFormatType.Solid
				tr.Fill.SolidColor.Color = Color.Black
			Next tr

			ppt.SaveToFile("AppendHTML_result.pptx", FileFormat.Pptx2010)
			Process.Start("AppendHTML_result.pptx")

		End Sub
	End Class
End Namespace
