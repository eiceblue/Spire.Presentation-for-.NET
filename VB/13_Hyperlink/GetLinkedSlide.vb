Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation.Charts

Namespace GetLinkedSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create Presentation
			Dim presentation As New Presentation()

			'Load ppt file
			presentation.LoadFromFile("..\..\..\..\..\..\Data\linkedSlide.pptx")

			'Get the second slide
			Dim slide As ISlide = presentation.Slides(1)

			'Get the first shape of the second slide
			Dim shape As IAutoShape = TryCast(slide.Shapes(0), IAutoShape)

			'Get the linked slide index
			If shape.Click.ActionType = HyperlinkActionType.GotoSlide Then
				Dim targetSlide As ISlide = shape.Click.TargetSlide
				MessageBox.Show("Linked slide number = " & targetSlide.SlideNumber)
			End If
		End Sub
	End Class
End Namespace