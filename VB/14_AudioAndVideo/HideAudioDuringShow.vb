Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation.Charts

Namespace HideAudioDuringShow
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create Presentation
			Dim presentation As New Presentation()

			'Load ppt file
			presentation.LoadFromFile("..\..\..\..\..\..\Data\audio.pptx")

			'Get the first slide
			Dim slide As ISlide=presentation.Slides(0)

			'Hide Audio during show
			For Each shape As Shape In slide.Shapes
				If TypeOf shape Is IAudio Then
					Dim audio As IAudio = TryCast(shape, IAudio)
					audio.HideAtShowing = True
				End If
			Next shape

			'Save the file
			Dim result As String = "HideAudioDuringShow_result.pptx"
			presentation.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013)

			'Launching the result file.
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