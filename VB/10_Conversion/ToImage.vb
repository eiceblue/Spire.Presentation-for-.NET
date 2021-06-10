Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ToImage.pptx")

			'Save PPT document to images
			For i As Integer = 0 To presentation.Slides.Count - 1
				Dim fileName As String = String.Format("ToImage-img-{0}.png", i)
				Dim image As Image = presentation.Slides(i).SaveAsImage()
				image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png)
				Process.Start(fileName)
			Next i

		End Sub
	End Class
End Namespace