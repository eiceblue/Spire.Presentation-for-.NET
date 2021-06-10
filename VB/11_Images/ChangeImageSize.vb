Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ChangeImageSize
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ExtractImage.pptx")

			'
			Dim scale As Single=0.5f
			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					If TypeOf shape Is IEmbedImage Then
						Dim image As IEmbedImage = TryCast(shape, IEmbedImage)
						image.Width = image.Width * scale
						image.Height = image.Height * scale
					End If
				Next shape
			Next slide

			'Save the document
			Dim result As String="ChangeImageSize_result.pptx"
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