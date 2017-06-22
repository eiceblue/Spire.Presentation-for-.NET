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
			'create PPT document
			Dim presentation As New Presentation()

			'load the PPT with password
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Transitions.pptx")

			'set the first slide transition as push and sound mode
			presentation.Slides(0).SlideShowTransition.Type = TransitionType.Push
			presentation.Slides(0).SlideShowTransition.SoundMode = TransitionSoundMode.StartSound

			'set the second slide transition as circle and set the speed 
			presentation.Slides(1).SlideShowTransition.Type = TransitionType.Fade
			presentation.Slides(1).SlideShowTransition.Speed = TransitionSpeed.Slow

			'save the file
			presentation.SaveToFile("setTransition.pptx", FileFormat.Pptx2010)
			Process.Start("setTransition.pptx")
		End Sub
	End Class
End Namespace
