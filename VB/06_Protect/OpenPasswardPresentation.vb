Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace OpenPasswardPresentation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create PPT document
			Dim presentation As New Presentation()

			'load the PPT with password
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Password.pptx", FileFormat.Pptx2010, textBox1.Text)

			'save as a new PPT with original password
			presentation.SaveToFile("New.pptx", FileFormat.Pptx2010)
			Process.Start("New.pptx")

		End Sub
	End Class
End Namespace
