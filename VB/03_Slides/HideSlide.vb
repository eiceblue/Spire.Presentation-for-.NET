Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace HideSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load PPT file from disk
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\HideSlide.pptx")

			'Hide the second slide
			ppt.Slides(1).Hidden = True

			'Save the document
			ppt.SaveToFile("HideSlide.pptx", FileFormat.Pptx2010)
			Process.Start("HideSlide.pptx")
		End Sub
	End Class
End Namespace
