Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace AddImageWatermark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_1.pptx")

			'Get the image you want to add as image watermark.
			Dim image As IImageData = presentation.Images.Append(Image.FromFile("..\..\..\..\..\..\Data\Logo.png"))

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            FileStream fileStream = new FileStream(@"..\..\..\..\..\..\Data\Logo.png", FileMode.Open, FileAccess.Read, FileShare.Read);
'            byte[] bytes = new byte[fileStream.Length];
'            fileStream.Read(bytes, 0, bytes.Length);
'            fileStream.Close();
'            Stream stream = new MemoryStream(bytes);
'            IImageData image = presentation.Images.Append(stream);
'            stream.Close();
'            fileStream.Close();
'            

			'Set the properties of SlideBackground, and then fill the image as watermark.
			presentation.Slides(0).SlideBackground.Type = Spire.Presentation.Drawing.BackgroundType.Custom
			presentation.Slides(0).SlideBackground.Fill.FillType = FillFormatType.Picture
			presentation.Slides(0).SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch
			presentation.Slides(0).SlideBackground.Fill.PictureFill.Picture.EmbedImage = image

			Dim result As String = "Result-AddImageWatermark.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file.
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				System.Diagnostics.Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace