Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace SaveChartAsImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document 
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SaveChartAsImage.pptx")

			'Save chart as image in .png format
			Dim image As Image = presentation.Slides(0).Shapes.SaveAsImage(0)
			image.Save("Chart_result.png", System.Drawing.Imaging.ImageFormat.Png)

			Process.Start("Chart_result.png")
		End Sub
	End Class
End Namespace
