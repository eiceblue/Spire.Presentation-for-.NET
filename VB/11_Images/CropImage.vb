Imports Spire.Presentation

Namespace CropImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\CropImage.pptx")

			'Get the first shape in first slide
			Dim shape As IShape = ppt.Slides(0).Shapes(0)

			'If the shape is SlidePicture
			If TypeOf shape Is SlidePicture Then
				Dim slidePicture As SlidePicture = CType(shape, SlidePicture)
				'Crop image
				slidePicture.Crop(slidePicture.Left + 50f, slidePicture.Top + 50f, 100f, 200f)
			End If

			'Save the document
			Dim result As String = "CropImage_out.pptx"
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
