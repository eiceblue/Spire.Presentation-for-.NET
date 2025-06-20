Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports Spire.Presentation.Drawing.Animation

Namespace AddExitAnimationForShape
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
			Dim starShape As IShape = slide.Shapes.AppendShape(ShapeType.FivePointedStar, New RectangleF(250, 100, 200, 200))
			starShape.Fill.FillType = FillFormatType.Solid
			starShape.Fill.SolidColor.KnownColor = KnownColors.LightBlue

			'Add random bars effect to the shape
			Dim effect As AnimationEffect = slide.Timeline.MainSequence.AddEffect(starShape, AnimationEffectType.RandomBars)

			'Change effect type from entrance to exit
			effect.PresetClassType = TimeNodePresetClassType.Exit

			'Save the document
			Dim result As String = "AddExitAnimationForShape.pptx"
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