Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace RotateShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create PPT document 
			Dim presentation As New Presentation()

			'append new shape - Triangle
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Triangle, New RectangleF(100, 100, 100, 100))
			'set rotation to 180
			shape.Rotation = 180

			'set the color and fill style of shape
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.BlueViolet
			shape.ShapeStyle.LineColor.Color = Color.Black

			'save the document
			presentation.SaveToFile("RotateShape.pptx", FileFormat.Pptx2010)
			Process.Start("RotateShape.pptx")
		End Sub

		Private Sub lblDescription_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lblDescription.Click

		End Sub
	End Class
End Namespace
