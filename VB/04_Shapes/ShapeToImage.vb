Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace ShapeToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ShapeToImage.pptx")

			For i As Integer = 0 To presentation.Slides(0).Shapes.Count - 1
				Dim fileName As String = String.Format("Picture-{0}.png", i)
				'Save shapes as images
				Dim image As Image = presentation.Slides(0).Shapes.SaveAsImage(i)
				image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png)
				Process.Start(fileName)
			Next i
		End Sub
	End Class
End Namespace
