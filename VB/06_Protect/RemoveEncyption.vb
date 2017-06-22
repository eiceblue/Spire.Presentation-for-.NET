Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace RemoveEncyption
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create PPT document
			Dim presentation As New Presentation()

			'load the PPT with password
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Password.pptx", FileFormat.Pptx2010, "123456")

			'remove encryption
			presentation.RemoveEncryption()

			'save the document
			presentation.SaveToFile("RemoveEncryption.pptx", FileFormat.Pptx2010)
			Process.Start("RemoveEncryption.pptx")
		End Sub
	End Class
End Namespace
