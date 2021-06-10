Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace HeaderAndFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			presentation.LoadFromFile("..\..\..\..\..\..\Data\HeaderAndFooter.pptx")

			'Add footer
			presentation.SetFooterText("Demo of Spire.Presentation")

			'Set the footer visible
			presentation.FooterVisible = True

			'Set the page number visible
			presentation.SlideNumberVisible = True

			'Set the date visible
			presentation.DateTimeVisible = True

			'Save the document
			presentation.SaveToFile("HeaderAndFooter_result.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("HeaderAndFooter_result.pptx")
		End Sub
	End Class
End Namespace