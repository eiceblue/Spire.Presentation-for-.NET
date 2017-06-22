Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace RemoveSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\source.pptx")

			'remove the second slide
			presentation.Slides.RemoveAt(1)

			presentation.SaveToFile("RemovedSlide.pptx", FileFormat.Pptx2010)
			Process.Start("RemovedSlide.pptx")
		End Sub
	End Class
End Namespace
