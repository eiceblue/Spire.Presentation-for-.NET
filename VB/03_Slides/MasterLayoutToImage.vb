Imports Spire.Presentation

Namespace MasterLayoutToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load the file 
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\CloneMaster2.pptx")

			' Iterate the masters
			For i As Integer = 0 To ppt.Masters(0).Layouts.Count - 1
				' Save layouts as images
				Dim image As Image = ppt.Masters(0).Layouts(i).SaveAsImage()
				Dim fileName As String = String.Format("{0}.png", i)
				image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png)
			Next i

			ppt.Dispose()
		End Sub
	End Class
End Namespace