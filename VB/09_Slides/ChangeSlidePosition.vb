Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace ChangeSlidePosition
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create PPT document
			Dim presentation As New Presentation()

			'load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ReorderSlidePosition.pptx")

			'move the first slide to the second slide position
			Dim slide As ISlide = presentation.Slides(0)
			slide.SlideNumber = 2

			'save the document
			presentation.SaveToFile("ChangedPosition.pptx", FileFormat.Pptx2010)
			Process.Start("ChangedPosition.pptx")
		End Sub
	End Class
End Namespace
