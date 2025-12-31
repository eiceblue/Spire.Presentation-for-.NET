Imports System
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace ExtractImageFromSpecificSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\Images.pptx")

			'Get the pictures on the second slide and save them to image file
			Dim i As Integer = 0
			'Traverse all shapes in the second slide
			For Each s As IShape In ppt.Slides(1).Shapes
				'It is the SlidePicture object
				If TypeOf s Is SlidePicture Then
					'Save to image
					Dim ps As SlidePicture = TryCast(s, SlidePicture)
					ps.PictureFill.Picture.EmbedImage.Image.Save(String.Format("{0}.png", i))

					'////////////////Use the following code for netstandard dlls/////////////////////////
'                    
'                    String fileName = string.Format("SlidePic_{0}.png", i);
'                    SkiaSharp.SKImage image = SkiaSharp.SKImage.FromBitmap(ps.PictureFill.Picture.EmbedImage.Image);
'                    FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
'                    image.Encode().SaveTo(fileStream);
'                    fileStream.Flush();
'                    image.Dispose();
'                    

					i += 1
				End If
				'It is the PictureShape object
				If TypeOf s Is PictureShape Then
					'Save to image
					Dim ps As PictureShape = TryCast(s, PictureShape)
					ps.EmbedImage.Image.Save(String.Format("{0}.png", i))

					'////////////////Use the following code for netstandard dlls/////////////////////////
'                    
'                    String fileName = string.Format("ShapePic_{0}.png", i);
'                    SkiaSharp.SKImage image = SkiaSharp.SKImage.FromBitmap(ps.EmbedImage.Image);
'                    FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
'                    image.Encode().SaveTo(fileStream);
'                    fileStream.Flush();
'                    image.Dispose();
'                    
					i += 1
				End If
			Next s
		End Sub
	End Class
End Namespace