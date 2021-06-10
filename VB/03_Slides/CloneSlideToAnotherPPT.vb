Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace CloneSlideToAnotherPPT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\CloneSlideToAnotherPPT-2.pptx")

			'Load the another document and choose the first slide to be cloned
			Dim ppt1 As New Presentation()
			ppt1.LoadFromFile("..\..\..\..\..\..\Data\CloneSlideToAnotherPPT-1.pptx")
			Dim slide1 As ISlide = ppt1.Slides(0)

			'Insert the slide to the specified index in the source presentation
			Dim index As Integer = 1
			presentation.Slides.Insert(index, slide1)

			'Save the document
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("Output.pptx")
		End Sub
	End Class
End Namespace
