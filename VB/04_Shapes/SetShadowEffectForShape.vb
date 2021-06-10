Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetShadowEffectForShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			Dim slide As ISlide = ppt.Slides(0)

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			slide.Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Add a shape to slide.
			Dim rect1 As New RectangleF(200, 150, 300, 120)
			Dim shape As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect1)
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.LightBlue
			shape.Line.FillType = FillFormatType.None
			shape.TextFrame.Text = "This demo shows how to apply shadow effect to shape."
			shape.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
			shape.TextFrame.TextRange.Fill.SolidColor.Color = Color.Black

			'Create an inner shadow effect through InnerShadowEffect object. 
			Dim innerShadow As New InnerShadowEffect()
			innerShadow.BlurRadius = 20
			innerShadow.Direction = 0
			innerShadow.Distance = 0
			innerShadow.ColorFormat.Color = Color.Black

			'Apply the shadow effect to shape.
			shape.EffectDag.InnerShadowEffect = innerShadow

			'Save the document
			Dim result As String = "SetShadowEffectForShape.pptx"
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