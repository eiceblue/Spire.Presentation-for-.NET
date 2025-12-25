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
			'Create PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SetBackground.pptx")

			'Set the background of the first slide to Gradient color
			presentation.Slides(0).SlideBackground.Type = BackgroundType.Custom
			presentation.Slides(0).SlideBackground.Fill.FillType = FillFormatType.Gradient
			presentation.Slides(0).SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear
			presentation.Slides(0).SlideBackground.Fill.Gradient.GradientStyle = Spire.Presentation.Drawing.GradientStyle.FromTopLeftCorner
			presentation.Slides(0).SlideBackground.Fill.Gradient.GradientStops.Append(1f, KnownColors.SkyBlue)
			presentation.Slides(0).SlideBackground.Fill.Gradient.GradientStops.Append(0f, KnownColors.White)

			'Set the background of the second slide to Solid color
			presentation.Slides(1).SlideBackground.Type = BackgroundType.Custom
			presentation.Slides(1).SlideBackground.Fill.FillType = FillFormatType.Solid
			presentation.Slides(1).SlideBackground.Fill.SolidColor.Color = Color.SkyBlue

			presentation.Slides.Append()
			'Set the background of the third slide to picture
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(2).SlideBackground.Fill.FillType = FillFormatType.Picture
			Dim image As IEmbedImage = presentation.Slides(2).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			presentation.Slides(2).SlideBackground.Fill.PictureFill.Picture.EmbedImage = TryCast(image, IImageData)


			'Save the document
			presentation.SaveToFile("SetBackground.pptx", FileFormat.Pptx2010)
			Process.Start("SetBackground.pptx")
		End Sub
	End Class
End Namespace
