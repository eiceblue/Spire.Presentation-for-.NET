Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetOutlineAndEffectForShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Get the first slide
			Dim slide As ISlide = ppt.Slides(0)

			'Set background Image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			slide.Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Draw a Rectangle shape
			Dim shape As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(150, 180, 100, 50))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.SkyBlue
			'Set outline color
			shape.ShapeStyle.LineColor.Color = Color.Red
			'Set shadow effect
			Dim shadow As New PresetShadow()
			shadow.ColorFormat.Color = Color.LightSkyBlue
			shadow.Preset = PresetShadowValue.FrontRightPerspective
			shadow.Distance = 10.0
			shadow.Direction = 225.0f
			shape.EffectDag.PresetShadowEffect = shadow

			'Draw a Ellipse shape
			shape = slide.Shapes.AppendShape(ShapeType.Ellipse, New RectangleF(400, 150, 100, 100))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.SkyBlue
			'Set outline color
			shape.ShapeStyle.LineColor.Color = Color.Yellow
			'Set shadow effect
			Dim glow As New GlowEffect()
			glow.ColorFormat.Color = Color.LightPink
			glow.Radius = 20.0
			shape.EffectDag.GlowEffect = glow

			'Save the document
			Dim result As String = "SetOutlineAndEffectForShape.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace