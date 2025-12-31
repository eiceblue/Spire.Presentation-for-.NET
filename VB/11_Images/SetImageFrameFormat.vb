Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetImageFrameFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load an image
			Dim imageFile As String = "..\..\..\..\..\..\Data\iceblueLogo.png"
			Dim image As Image = Image.FromFile(imageFile)

			'Add the image in document
			Dim imageData As IImageData = presentation.Images.Append(image)

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            FileStream fileStream = new FileStream(imageFile, FileMode.Open, FileAccess.Read, FileShare.Read);
'            byte[] bytes = new byte[fileStream.Length];
'            fileStream.Read(bytes, 0, bytes.Length);
'            fileStream.Close();
'            Stream stream = new MemoryStream(bytes);
'            IImageData imageData = presentation.Images.Append(stream);
'            stream.Close();
'            fileStream.Close();
'            

			Dim rect As New RectangleF(100,100,imageData.Width\2,imageData.Height\2)
			Dim pptImage As IEmbedImage = presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, imageData, rect)

			'Set the formatting of the image frame
			pptImage.Line.FillFormat.FillType = FillFormatType.Solid
			pptImage.Line.FillFormat.SolidFillColor.Color = Color.LightBlue
			pptImage.Line.Width = 5
			pptImage.Rotation = -45

			'Save the document
			Dim result As String = "SetImageFrameFormat_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				System.Diagnostics.Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace