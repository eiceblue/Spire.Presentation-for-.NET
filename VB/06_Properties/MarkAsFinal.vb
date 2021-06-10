Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace MarkAsFinal
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\MarkAsFinal.pptx")

			'Mark the document as final
			presentation.DocumentProperty("_MarkAsFinal") = True

			'Save the document
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("Output.pptx")
		End Sub
	End Class
End Namespace
