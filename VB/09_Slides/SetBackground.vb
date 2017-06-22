Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace SetBackground
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create PPT document
			Dim presentation As New Presentation()

			'add new slide
			presentation.Slides.Append()
			'add new slide
			presentation.Slides.Append()

			'set the background of the first slide to Gradient color
			presentation.Slides(0).SlideBackground.Type = BackgroundType.Custom
			presentation.Slides(0).SlideBackground.Fill.FillType = FillFormatType.Gradient
			presentation.Slides(0).SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear
			presentation.Slides(0).SlideBackground.Fill.Gradient.GradientStyle = Spire.Presentation.Drawing.GradientStyle.FromCorner1
			presentation.Slides(0).SlideBackground.Fill.Gradient.GradientStops.Append(1f, KnownColors.LightGreen)
			presentation.Slides(0).SlideBackground.Fill.Gradient.GradientStops.Append(0f, KnownColors.White)

			'set the background of the second slide to Solid color
			presentation.Slides(1).SlideBackground.Type = BackgroundType.Custom
			presentation.Slides(1).SlideBackground.Fill.FillType = FillFormatType.Solid
			presentation.Slides(1).SlideBackground.Fill.SolidColor.Color = Color.DarkSeaGreen

			'set the background of the third slide to picture
			Dim ImageFile As String = "..\..\..\..\..\..\Data\background.png"
			Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(2).SlideBackground.Fill.FillType = FillFormatType.Picture
			Dim image As IEmbedImage = presentation.Slides(2).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			presentation.Slides(2).SlideBackground.Fill.PictureFill.Picture.EmbedImage = TryCast(image, IImageData)

			'add shape and fill it with text in slides
			Dim shape As IAutoShape
			Dim textRange As TextRange
			For i As Integer = 0 To 2
				shape = presentation.Slides(i).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 70, 600, 100))
				shape.ShapeStyle.LineColor.Color = Color.White
				shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

				shape.AppendTextFrame("Demonstrates to how to set the background style of slides.")

				'set the Font and fill style
				textRange = shape.TextFrame.TextRange
				textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
				textRange.Fill.SolidColor.Color = Color.Black
				textRange.LatinFont = New TextFont("Arial Black")
			Next i
			'save the document
			presentation.SaveToFile("background.pptx", FileFormat.Pptx2010)
			Process.Start("background.pptx")
		End Sub
	End Class
End Namespace
