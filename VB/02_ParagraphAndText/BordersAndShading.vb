Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace BordersAndShading
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\BordersAndShading.pptx")

			Dim shape As IAutoShape = CType(presentation.Slides(0).Shapes(0), IAutoShape)

			'Set line color and width of the border
			shape.Line.FillType = FillFormatType.Solid
			shape.Line.Width = 3
			shape.Line.SolidFillColor.Color = Color.LightYellow

			'Set the gradient fill color of shape
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Gradient
			shape.Fill.Gradient.GradientShape = Spire.Presentation.Drawing.GradientShapeType.Linear
			shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.LightBlue)
			shape.Fill.Gradient.GradientStops.Append(0, KnownColors.LightSkyBlue)

			'Set the shadow for the shape
			Dim shadow As New Spire.Presentation.Drawing.OuterShadowEffect()
			shadow.BlurRadius = 20
			shadow.Direction = 30
			shadow.Distance = 8
			shadow.ColorFormat.Color = Color.LightSeaGreen
			shape.EffectDag.OuterShadowEffect = shadow

			'Save the document
			presentation.SaveToFile("BordersAndShading_result.pptx", FileFormat.Pptx2007)
			Process.Start("BordersAndShading_result.pptx")
		End Sub
	End Class
End Namespace