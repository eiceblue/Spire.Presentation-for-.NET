Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace CreateSlideMasterAndApply
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			ppt.SlideSize.Type = SlideSizeType.Screen16x9

			'Add slides
			For i As Integer = 0 To 3
				ppt.Slides.Append()
			Next i

			'Get the first default slide master
			Dim first_master As IMasterSlide = ppt.Masters(0)

			'Append another slide master
			ppt.Masters.AppendSlide(first_master)
			Dim second_master As IMasterSlide = ppt.Masters(1)

			'Set different background image for the two slide masters
			Dim pic1 As String = "..\..\..\..\..\..\Data\bg.png"
			Dim pic2 As String = "..\..\..\..\..\..\Data\Setbackground.png"
			'The first slide master
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			first_master.SlideBackground.Fill.FillType = FillFormatType.Picture
			Dim image1 As IEmbedImage = first_master.Shapes.AppendEmbedImage(ShapeType.Rectangle, pic1, rect)
			first_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = TryCast(image1, IImageData)
			'The second slide master
			second_master.SlideBackground.Fill.FillType = FillFormatType.Picture
			Dim image2 As IEmbedImage = second_master.Shapes.AppendEmbedImage(ShapeType.Rectangle, pic2, rect)
			second_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = TryCast(image2, IImageData)

			'Apply the first master with layout to the first slide
			ppt.Slides(0).Layout = first_master.Layouts(1)

			'Apply the second master with layout to other slides
			For i As Integer = 1 To ppt.Slides.Count - 1
				ppt.Slides(i).Layout = second_master.Layouts(8)
			Next i

			'Save the document
			Dim result As String = "CreateSlideMasterAndApply.pptx"
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