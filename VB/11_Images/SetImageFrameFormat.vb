Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetImageFrameFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load an image
			Dim imageFile As String = "..\..\..\..\..\..\Data\iceblueLogo.png"
			Dim image As Image = Image.FromFile(imageFile)

			'Add the image in document
			Dim imageData As IImageData = presentation.Images.Append(image)
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
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace