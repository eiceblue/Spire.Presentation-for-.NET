Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace AddShapes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Set background Image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Append new shape - Triangle and set style
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Triangle, New RectangleF(115, 130, 100, 100))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.LightGreen
			shape.ShapeStyle.LineColor.Color = Color.White

			'Append new shape - Ellipse
			shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Ellipse, New RectangleF(290, 130, 150, 100))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.LightSkyBlue
			shape.ShapeStyle.LineColor.Color = Color.White

			'Append new shape - Heart
			shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Heart, New RectangleF(470, 130, 130, 100))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.Red
			shape.ShapeStyle.LineColor.Color = Color.LightGray


			'Append new shape - FivePointedStar
			shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.FivePointedStar, New RectangleF(90, 270, 150, 150))
			shape.Fill.FillType = FillFormatType.Gradient
			shape.Fill.SolidColor.Color = Color.Black
			shape.ShapeStyle.LineColor.Color = Color.White

			'Append new shape - Rectangle
			shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(320, 290, 100, 120))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.Pink
			shape.ShapeStyle.LineColor.Color = Color.LightGray

			'Append new shape - BentUpArrow
			shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.BentUpArrow, New RectangleF(470, 300, 150, 100))

			'Set the color of shape
			shape.Fill.FillType = FillFormatType.Gradient
			shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.Olive)
			shape.Fill.Gradient.GradientStops.Append(0, KnownColors.PowderBlue)
			shape.ShapeStyle.LineColor.Color = Color.White

			'Save the document
			presentation.SaveToFile("AddShapes_result.pptx", FileFormat.Pptx2010)
			Process.Start("AddShapes_result.pptx")
		End Sub
	End Class
End Namespace