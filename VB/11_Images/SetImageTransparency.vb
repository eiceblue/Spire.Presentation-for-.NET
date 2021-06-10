Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetImageTransparency
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Create an Image from the specified file
			Dim imagePath As String = "..\..\..\..\..\..\Data\Logo.png"
			Dim image As Image = Image.FromFile(imagePath)
			Dim width As Single = image.Width
			Dim height As Single = image.Height
			Dim rect1 As New RectangleF(200, 100, width, height)
			'Add a shape
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rect1)
			shape.Line.FillType = FillFormatType.None
			'Fill shape with image
			shape.Fill.FillType = FillFormatType.Picture
			shape.Fill.PictureFill.Picture.Url = imagePath
			shape.Fill.PictureFill.FillType = PictureFillType.Stretch
			'Set transparency on image
			shape.Fill.PictureFill.Picture.Transparency = 50

			'Save the document
			Dim result As String = "SetImageTransparency.pptx"
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