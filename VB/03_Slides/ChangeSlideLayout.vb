Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ChangeSlideLayout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\ChangeSlideLayout.pptx")

			'Change the layout of slide
			presentation.Slides(1).Layout = presentation.Masters(0).Layouts(4)

			'Save the document
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("Output.pptx")

		End Sub
	End Class
End Namespace
