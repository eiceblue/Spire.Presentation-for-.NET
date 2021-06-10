Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace ExtractAudio
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim loadPath As String = "..\..\..\..\..\..\Data\audio.pptx"
			Dim outPath As String = "extrctAudio.wav"
			Dim AudioData() As Byte = Nothing

			'Load a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile(loadPath)

			For Each shape As Shape In presentation.Slides(0).Shapes
				If TypeOf shape Is IAudio Then
					Dim audio As IAudio = TryCast(shape, IAudio)
					AudioData = audio.Data.Data
				End If
			Next shape

			Using fs As New FileStream(outPath, FileMode.Create, FileAccess.Write)
				fs.Write(AudioData, 0, AudioData.Length)

			End Using
			Process.Start(outPath)
		End Sub
	End Class
End Namespace