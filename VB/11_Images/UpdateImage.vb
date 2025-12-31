Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace UpdateImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\UpdateImage.pptx")

			'Get the first slide
			Dim slide As ISlide = ppt.Slides(0)

			'Append a new image to replace an existing image
			Dim image As IImageData = ppt.Images.Append(Image.FromFile("..\..\..\..\..\..\Data\PresentationIcon.png"))

			'////////////////Use the following code for netstandard dlls/////////////////////////
'			
'			FileStream fileStream = new FileStream(@"..\..\..\..\..\..\Data\PresentationIcon.png", FileMode.Open, FileAccess.Read, FileShare.Read);
'            byte[] bytes = new byte[fileStream.Length];
'            fileStream.Read(bytes, 0, bytes.Length);
'            fileStream.Close();
'            Stream stream = new MemoryStream(bytes);
'            IImageData image = ppt.Images.Append(stream);
'            stream.Close();
'            fileStream.Close();
'			

			'Replace the image which title is "image1" with the new image
			For Each shape As IShape In slide.Shapes
				If TypeOf shape Is SlidePicture Then
					If shape.AlternativeTitle = "image1" Then
						TryCast(shape, SlidePicture).PictureFill.Picture.EmbedImage = image
					End If
				End If
			Next shape

			'Save the document
			Dim result As String = "UpdateImage.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				System.Diagnostics.Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace