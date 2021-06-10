Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.IO

Namespace ModifyOLEData
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ModifyOLEData.pptx")

			'Loop through the slides and shapes
			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					If TypeOf shape Is IOleObject Then
						'Find OLE object
						Dim oleObject As IOleObject = TryCast(shape, IOleObject)

						'Get its data and write to file
						Dim bytes() As Byte = oleObject.Data
						Dim pptStream As New MemoryStream(bytes)
						Dim stream As New MemoryStream()
						If oleObject.ProgId = "PowerPoint.Show.12" Then
							'Load the PPT stream
							Dim ppt As New Presentation()
							ppt.LoadFromStream(pptStream, Spire.Presentation.FileFormat.Auto)
							'Append an image in slide
							ppt.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, "..\..\..\..\..\..\Data\Logo.png", New RectangleF(50, 50, 100, 100))
							ppt.SaveToFile(stream, Spire.Presentation.FileFormat.Pptx2013)
							stream.Position = 0
							'Modify the data
							oleObject.Data = stream.ToArray()
						End If
					End If
				Next shape
			Next slide

			'Save the document
			Dim result As String = "ModifyOLEData_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace