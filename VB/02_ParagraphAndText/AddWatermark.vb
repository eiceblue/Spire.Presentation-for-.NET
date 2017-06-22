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
			'create a PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\table.pptx")

			'get the size of watermark string
			Dim gc As Graphics = Me.CreateGraphics()
			Dim size As SizeF = gc.MeasureString("E-iceblue", New Font("Arial", 45))

			'define a rectangle range
			Dim rect As New RectangleF((presentation.SlideSize.Size.Width - size.Width) / 2, (presentation.SlideSize.Size.Height - size.Height) / 2, size.Width, size.Height)

			'add a rectangle shape with a defined range
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(Spire.Presentation.ShapeType.Rectangle, rect)

			'set the style of shape
			shape.Fill.FillType = FillFormatType.None
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Rotation = -45
			shape.Locking.SelectionProtection = True
			shape.Line.FillType = FillFormatType.None

			'add text to shape
			shape.TextFrame.Text = "E-iceblue"
			Dim textRange As TextRange = shape.TextFrame.TextRange
			'set the style of the text range
			textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.FromArgb(120, Color.Black)
			textRange.FontHeight = 45

			presentation.SaveToFile("Watermark.pptx", FileFormat.Pptx2010)
			Process.Start("Watermark.pptx")
		End Sub
	End Class
End Namespace
