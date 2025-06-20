Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Animation

Namespace GetAnimationEffectInfo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\Animation.pptx")

			Dim sb As New StringBuilder()
			'Travel each slide
			For Each slide As ISlide In presentation.Slides
				For Each effect As AnimationEffect In slide.Timeline.MainSequence
					'Get the animation effect type
					Dim animationEffectType As AnimationEffectType = effect.AnimationEffectType
					sb.AppendLine("animation effect type:" & animationEffectType)

					'Get the slide number where the animation is located
					Dim slideNumber As Integer = slide.SlideNumber
					sb.AppendLine("slide number:" & slideNumber)

					'Get the shape name
					Dim shapeName As String = effect.ShapeTarget.Name
					sb.AppendLine("shape name:" & shapeName & vbLf)

				Next effect
			Next slide

			'Save the information of animation effect
			Dim result As String = "AnimationEffectInfo.txt"
			File.WriteAllText(result, sb.ToString())

			Process.Start(result)
		End Sub
	End Class
End Namespace