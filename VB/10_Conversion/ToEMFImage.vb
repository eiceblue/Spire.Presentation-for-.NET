Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace ToEMFImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ToEMFImage.pptx")

			'Save to EMF image
			presentation.Slides(0).SaveAsEMF("ToEMFImage.emf")
			Process.Start("ToEMFImage.emf")
		End Sub
	End Class
End Namespace
