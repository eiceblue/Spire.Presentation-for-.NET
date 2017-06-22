Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace FillShapeWithPicture
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create PPT document
			Dim ppt As New Presentation()

			'add a rectangle to the slide
			Dim shape As IAutoShape = CType(ppt.Slides(0).Shapes.AppendShape(ShapeType.DoubleWave, New RectangleF(100, 100, 400, 200)), IAutoShape)

			'fill the shape with picture
			Dim picUrl As String = "..\..\..\..\..\..\Data\bg.png"
			shape.Fill.FillType = FillFormatType.Picture
			shape.Fill.PictureFill.Picture.Url = picUrl
			shape.Fill.PictureFill.FillType = PictureFillType.Stretch
			shape.ShapeStyle.LineColor.Color = Color.Transparent

			ppt.SaveToFile("FillShapeWithPicture.pptx", FileFormat.Pptx2010)
			Process.Start("FillShapeWithPicture.pptx")
		End Sub
	End Class
End Namespace
