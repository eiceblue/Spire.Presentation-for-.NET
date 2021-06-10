Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace Encrypt
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Encrypt.pptx")

			'Get the password that the user entered
			Dim password As String = Me.textBox1.Text

			'Encrypy the document with the password
			presentation.Encrypt(password)

			'Save the document
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("Output.pptx")
		End Sub
	End Class
End Namespace