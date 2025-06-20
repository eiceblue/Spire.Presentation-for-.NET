Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Namespace BetterSlideTransitions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\SetTransitions.pptx")

			'Set the first slide transition as circle
			presentation.Slides(0).SlideShowTransition.Type = TransitionType.Circle

			' Set the transition time of 3 seconds
			presentation.Slides(0).SlideShowTransition.AdvanceOnClick = True
			presentation.Slides(0).SlideShowTransition.AdvanceAfterTime = 3000

			'Set the second slide transition as comb and set the speed 
			presentation.Slides(1).SlideShowTransition.Type = TransitionType.Comb
			presentation.Slides(1).SlideShowTransition.Speed = TransitionSpeed.Slow

			' Set the transition time of 5 seconds
			presentation.Slides(1).SlideShowTransition.AdvanceOnClick = True
			presentation.Slides(1).SlideShowTransition.AdvanceAfterTime = 5000

			' Set the third slide transition as zoom
			presentation.Slides(2).SlideShowTransition.Type = TransitionType.Zoom

			' Set the transition time of 7 seconds
			presentation.Slides(2).SlideShowTransition.AdvanceOnClick = True
			presentation.Slides(2).SlideShowTransition.AdvanceAfterTime = 7000


			Dim result As String = "BetterSlideTransitions_result.pptx"
			'Save the file
			presentation.SaveToFile(result, FileFormat.Pptx2010)

			Viewer(result)
		End Sub

		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace