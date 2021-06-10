Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel


Namespace SetRadiusOfRoundedRectangle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Insert a rounded rectangle and set its radious
			presentation.Slides(0).Shapes.InsertRoundRectangle(0, 160, 180, 100, 200, 10)

			'Append a rounded rectangle and set its radius
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendRoundRectangle(380, 180, 100, 200, 100)
			'Set the color and fill style of shape
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.SeaGreen
			shape.ShapeStyle.LineColor.Color = Color.White

			'Rotate the shape to 90 degree
			shape.Rotation = 90

			'Save the document to Pptx file
			Dim result As String = "SetRadiusOfRoundedRectangle.pptx"
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