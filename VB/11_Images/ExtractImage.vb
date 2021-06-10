Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace ExtractImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ExtractImage.pptx")

			For i As Integer = 0 To ppt.Images.Count - 1
				Dim ImageName As String = String.Format("Images-{0}.png", i)
				'Extract image
				Dim image As Image = ppt.Images(i).Image
				image.Save(ImageName)
				Process.Start(ImageName)
			Next i
		End Sub
	End Class
End Namespace
