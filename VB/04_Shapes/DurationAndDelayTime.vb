Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Collections
Imports Spire.Presentation.Drawing.Animation

Namespace DurationAndDelayTime
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\Animation.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			Dim animations As AnimationEffectCollection = slide.Timeline.MainSequence

			'Get duration time of animation
			Dim durationTime As Single = animations(0).Timing.Duration

			'Set new duration time of animation
			animations(0).Timing.Duration = 0.8f

			'Get delay time of animation
			Dim delayTime As Single = animations(0).Timing.TriggerDelayTime

			'Set new delay time of animation
			animations(0).Timing.TriggerDelayTime = 0.6f
			Dim result As String = "DurationAndDelayTime_result.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file.
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace