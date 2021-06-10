Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace FillShapeWithGradient
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim ppt As New Presentation()

			ppt.LoadFromFile("..\..\..\..\..\..\Data\FillShapeWithGradient.pptx")

			'Get the first shape and set the style to be Gradient
			Dim GradientShape As IAutoShape = TryCast(ppt.Slides(0).Shapes(0), IAutoShape)
			GradientShape.Fill.FillType = FillFormatType.Gradient
			GradientShape.Fill.Gradient.GradientStops.Append(0, Color.LightSkyBlue)
			GradientShape.Fill.Gradient.GradientStops.Append(1, Color.LightGray)

			ppt.SaveToFile("FillShapeWithGradient_result.pptx", FileFormat.Pptx2010)
			Process.Start("FillShapeWithGradient_result.pptx")
		End Sub

	End Class
End Namespace
