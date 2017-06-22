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
			'create PPT document and load PPT file from disk
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\source.pptx")

			'Load the another document and choose the first slide to be cloned.
			Dim ppt1 As New Presentation()
			ppt1.LoadFromFile("..\..\..\..\..\..\Data\Presentation1.pptx")
			Dim slide1 As ISlide = ppt1.Slides(0)

			'Insert the slide to the specified index in the source presentation
			Dim index As Integer = 1
			presentation.Slides.Insert(index, slide1)

			'save the document
			presentation.SaveToFile("ClonedSlide.pptx", FileFormat.Pptx2010)
			Process.Start("ClonedSlide.pptx")
		End Sub
	End Class
End Namespace
