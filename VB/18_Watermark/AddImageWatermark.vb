Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace AddImageWatermark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_1.pptx")

            'Get the image you want to add as image watermark.
            Dim imagedata As IImageData
            imagedata = presentation.Images.Append(Image.FromFile("..\..\..\..\..\..\Data\Logo.png"))

            'Set the properties of SlideBackground, and then fill the image as watermark.
            presentation.Slides(0).SlideBackground.Type = Spire.Presentation.Drawing.BackgroundType.Custom
			presentation.Slides(0).SlideBackground.Fill.FillType = FillFormatType.Picture
			presentation.Slides(0).SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch
            presentation.Slides(0).SlideBackground.Fill.PictureFill.Picture.EmbedImage = imagedata

            Dim result As String = "Result-AddImageWatermark.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file.
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace