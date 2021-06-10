Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace ApplySlideMaster
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\InputTemplate.pptx")

			'Get the first slide master from the presentation
			Dim masterSlide As IMasterSlide = ppt.Masters(0)

			'Customize the background of the slide master
			Dim backgroundPic As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			masterSlide.SlideBackground.Fill.FillType = FillFormatType.Picture
			Dim image As IEmbedImage = masterSlide.Shapes.AppendEmbedImage(ShapeType.Rectangle, backgroundPic, rect)
			masterSlide.SlideBackground.Fill.PictureFill.Picture.EmbedImage = TryCast(image, IImageData)

			'Change the color scheme
			masterSlide.Theme.ColorScheme.Accent1.Color = Color.Red
			masterSlide.Theme.ColorScheme.Accent2.Color = Color.RosyBrown
			masterSlide.Theme.ColorScheme.Accent3.Color = Color.Ivory
			masterSlide.Theme.ColorScheme.Accent4.Color = Color.Lavender
			masterSlide.Theme.ColorScheme.Accent5.Color = Color.Black

			'Save the document
			Dim result As String = "ApplySlideMaster.pptx"
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