Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ToPDF
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ToPDF.pptx")

			'Save the PPT to PDF file format
			presentation.SaveToFile("ToPdf.pdf", FileFormat.PDF)
			Process.Start("ToPdf.pdf")

		End Sub
	End Class
End Namespace