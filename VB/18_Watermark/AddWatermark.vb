Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace AddWatermark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AddWatermark.pptx")

			'Get the size of the watermark string
			Dim gc As Graphics = Me.CreateGraphics()
			Dim size As SizeF = gc.MeasureString("E-iceblue", New Font("Lucida Sans Unicode", 50))

			'Define a rectangle range
			Dim rect As New RectangleF((presentation.SlideSize.Size.Width - size.Width) / 2, (presentation.SlideSize.Size.Height - size.Height) / 2, size.Width, size.Height)

			'Add a rectangle shape with a defined range
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(Spire.Presentation.ShapeType.Rectangle, rect)

			'Set the style of the shape
			shape.Fill.FillType = FillFormatType.None
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Rotation = -45
			shape.Locking.SelectionProtection = True
			shape.Line.FillType = FillFormatType.None

			'Add text to the shape
			shape.TextFrame.Text = "E-iceblue"
			Dim textRange As TextRange = shape.TextFrame.TextRange
			'Set the style of the text range
			textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.FromArgb(120, Color.HotPink)
			textRange.FontHeight = 50

			'Save the document and launch
			presentation.SaveToFile("Watermark_result.pptx", FileFormat.Pptx2010)
			Process.Start("Watermark_result.pptx")
		End Sub
	End Class
End Namespace
