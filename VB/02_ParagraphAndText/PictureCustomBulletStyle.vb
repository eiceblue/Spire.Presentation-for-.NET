Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace PictureCustomBulletStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\Bullets.pptx")

			'Get the second shape on the first slide
			Dim shape As IAutoShape = TryCast(ppt.Slides(0).Shapes(1), IAutoShape)

			'Traverse through the paragraphs in the shape
			For Each paragraph As TextParagraph In shape.TextFrame.Paragraphs
				'Set the bullet style of paragraph as picture
				paragraph.BulletType = TextBulletType.Picture

				'////////////////Use the following code for netstandard dlls/////////////////////////
'                
'                FileStream fileStream = new FileStream(@"..\..\..\..\..\..\Data\icon.png", FileMode.Open, FileAccess.Read, FileShare.Read);
'                byte[] bytes = new byte[fileStream.Length];
'                fileStream.Read(bytes, 0, bytes.Length);
'                fileStream.Close();
'                Stream stream = new MemoryStream(bytes);
'                paragraph.BulletPicture.EmbedImage = ppt.Images.Append(stream);
'                stream.Close();
'                

				'Load a picture
				Dim bulletPicture As Image = Image.FromFile("..\..\..\..\..\..\Data\icon.png")
				'Add the picture as the bullet style of paragraph
				paragraph.BulletPicture.EmbedImage = ppt.Images.Append(bulletPicture)
			Next paragraph

			'Save the document
			Dim result As String = "PictureCustomBulletStyle.pptx"
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