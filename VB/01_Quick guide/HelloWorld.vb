Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace HelloWorld
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Add a new shape to the PPT document
			Dim rec As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 250, 80, 500, 150)
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)

			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

			'Add text to the shape
			shape.AppendTextFrame("Hello World!")

			'Set the font and fill style of the text
			Dim textRange As TextRange = shape.TextFrame.TextRange
			textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.CadetBlue
			textRange.FontHeight = 66
			textRange.LatinFont = New TextFont("Lucida Sans Unicode")

			'Save the document
			presentation.SaveToFile("HelloWorld.pptx", FileFormat.Pptx2010)
			Process.Start("HelloWorld.pptx")
		End Sub
	End Class
End Namespace