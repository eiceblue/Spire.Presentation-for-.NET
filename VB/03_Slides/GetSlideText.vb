Imports Spire.Presentation
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Text

Namespace GetSlideText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\GetSlideText.pptx")

			'Foreach the slide and get text
			For Each slide As ISlide In ppt.Slides
				Dim arrayList As ArrayList = slide.GetAllTextFrame()
				For Each text As String In arrayList
					MessageBox.Show(text)
				Next text
			Next slide
		End Sub
	End Class
End Namespace
