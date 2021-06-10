Imports System.IO
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.TimeLine

Namespace ObtainSoundEffect
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\Animation.pptx")

			'Get the first slide
			Dim slide As ISlide = ppt.Slides(0)

			'Get the audio in a time node
			Dim audio As TimeNodeAudio = slide.Timeline.MainSequence(0).TimeNodeAudios(0)

			'Get the properties of the audio, such as sound name, volume or detect if it's mute
			Dim sb As New StringBuilder()
			sb.AppendLine("SoundName: " & audio.SoundName)
			sb.AppendLine("Volume: " & audio.Volume)
			sb.AppendLine("IsMute: " & audio.IsMute)

			'Save the properties of the audio to Text file
			Dim result As String = "ObtainSoundEffect.txt"
			File.WriteAllText(result, sb.ToString())
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace