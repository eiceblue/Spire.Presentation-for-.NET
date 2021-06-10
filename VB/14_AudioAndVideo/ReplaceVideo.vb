Imports Spire.Presentation
Imports Spire.Presentation.Collections
Imports System.IO


Namespace ReplaceVideo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim ppt As New Presentation()

			'Load the PPT document from disk.
			ppt.LoadFromFile("..\..\..\..\..\..\Data\video.pptx")

			Dim videos As VideoCollection = ppt.Videos

			'Traverse all the slides of PPT file
			For Each sld As ISlide In ppt.Slides
				'Traverse all the shapes of slides
				For Each sp As Shape In sld.Shapes
					'If shape is IVideo
					If TypeOf sp Is IVideo Then
						'Replace the video
						Dim video As IVideo = TryCast(sp, IVideo)
						'Load the video document from disk.
						Dim bts() As Byte = File.ReadAllBytes("..\..\..\..\..\..\Data\repleaceVido.mp4")
						Dim videoData As VideoData = videos.Append(bts)
						video.EmbeddedVideoData = videoData
					End If
				Next sp
			Next sld

			'Save the document
			Dim outputFile As String = "replaceVideo.pptx"
			ppt.SaveToFile(outputFile, FileFormat.Pptx2013)

			'Launch the PPT file
			FileViewer(outputFile)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
