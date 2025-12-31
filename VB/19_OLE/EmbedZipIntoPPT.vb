Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace EmbedZipIntoPPT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Load a ppt document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\EmbedZipIntoPPT.pptx")

			'Load a zip object
			Dim zipPath As String = "..\..\..\..\..\..\Data\test.zip"
			Dim data() As Byte = File.ReadAllBytes(zipPath)

			Dim rec As New Rectangle(80, 60, 100, 100)

			'Insert the zip object to presentation
			Dim ole As IOleObject = ppt.Slides(0).Shapes.AppendOleObject("test.zip", data, rec)
			ole.ProgId = "Package"
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\icon.png")
			Dim oleImage As IImageData = ppt.Images.Append(image)

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            FileStream stream = new FileStream(@"..\..\..\..\..\..\Data\icon.png", FileMode.Open);
'            IImageData oleImage = ppt.Images.Append(stream);
'            

			ole.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage

			'Save the document
			ppt.SaveToFile("EmbedZipIntoPPT_result.pptx", FileFormat.Pptx2010)

			'Launch the file
			OutputViewer("EmbedZipIntoPPT_result.pptx")
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				System.Diagnostics.Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace