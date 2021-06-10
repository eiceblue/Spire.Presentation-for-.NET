Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Collections
Imports Spire.Presentation.Drawing.Animation

Namespace RemoveImages
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\RemoveImages.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			For i As Integer = slide.Shapes.Count-1 To 0 Step -1
				'It is the SlidePicture object
				If TypeOf slide.Shapes(i) Is SlidePicture Then
					slide.Shapes.RemoveAt(i)
				End If
			Next i

			Dim result As String = "RemoveImages_result.pptx"

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