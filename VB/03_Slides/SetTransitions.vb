Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports System.ComponentModel
Imports System.Text

Namespace SetTransitions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT with password
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\SetTransitions.pptx")

			'Set the first slide transition as push and sound mode
			presentation.Slides(0).SlideShowTransition.Type = TransitionType.Push
			presentation.Slides(0).SlideShowTransition.SoundMode = TransitionSoundMode.StartSound

			'Set the second slide transition as circle and set the speed 
			presentation.Slides(1).SlideShowTransition.Type = TransitionType.Fade
			presentation.Slides(1).SlideShowTransition.Speed = TransitionSpeed.Slow

			'Save the file
			presentation.SaveToFile("SetTransition.pptx", FileFormat.Pptx2010)
			Process.Start("SetTransition.pptx")
		End Sub
	End Class
End Namespace
