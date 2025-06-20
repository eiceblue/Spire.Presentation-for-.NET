Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports Spire.Presentation.Drawing.Animation

Namespace ApplyAnimationOnText
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

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			slide.Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Add a shape to the slide
			Dim shape As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(250, 150, 200, 100))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.LightBlue
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.AppendTextFrame("This demo shows how to apply animation on text in PPT document.")

			'Apply animation to the text in shape
			Dim animation As AnimationEffect = shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Float)
			animation.SetStartEndParagraphs(0, 0)

			'Save the document
			Dim result As String = "ApplyAnimationOnText.pptx"
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