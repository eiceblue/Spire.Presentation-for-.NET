Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ToPPTX
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ToPPTX.ppt")

			'Save the PPT document to PPTX file format
			presentation.SaveToFile("ToPPTX.pptx", FileFormat.Pptx2010)
			Process.Start("ToPPTX.pptx")
		End Sub
	End Class
End Namespace