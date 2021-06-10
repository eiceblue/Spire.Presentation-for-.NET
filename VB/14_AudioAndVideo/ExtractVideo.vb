Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO

Namespace ExtractVideo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT document from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\video.pptx")

			'Define a variable 
			Dim i As Integer = 0

			'String for output file 
			Dim result As String = String.Format("Video{0}.avi", i)

			'Traverse all the slides of PPT file
			For Each slide As ISlide In presentation.Slides
				'Traverse all the shapes of slides
				For Each shape As IShape In slide.Shapes
					'If shape is IVideo
					If TypeOf shape Is IVideo Then
						'Save the video
						TryCast(shape, IVideo).EmbeddedVideoData.SaveToFile(result)
						i += 1
					End If
				Next shape
			Next slide
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