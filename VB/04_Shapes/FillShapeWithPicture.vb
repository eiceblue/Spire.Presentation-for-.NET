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
			'Load a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\FillShapeWithPicture.pptx")

			'Get the first shape and set the style to be Gradient
			Dim shape As IAutoShape = TryCast(ppt.Slides(0).Shapes(0), IAutoShape)

			'Fill the shape with picture
			Dim picUrl As String = "..\..\..\..\..\..\Data\backgroundImg.png"
			shape.Fill.FillType = FillFormatType.Picture
			shape.Fill.PictureFill.Picture.Url = picUrl
			shape.Fill.PictureFill.FillType = PictureFillType.Stretch

			ppt.SaveToFile("FillShapeWithPicture_result.pptx", FileFormat.Pptx2010)
			Process.Start("FillShapeWithPicture_result.pptx")
		End Sub
	End Class
End Namespace
