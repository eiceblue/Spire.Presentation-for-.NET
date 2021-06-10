Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace OpenEncryptedPPT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the PPT with password
			presentation.LoadFromFile("..\..\..\..\..\..\Data\OpenEncryptedPPT.pptx", FileFormat.Pptx2010, textBox1.Text)

			'Save as a new PPT with original password
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("Output.pptx")

		End Sub
	End Class
End Namespace
