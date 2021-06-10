Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace RemoveTextOrImageWatermark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\RemoveTextAndImageWatermarks.pptx")

			'Remove text watermark by removing the shape which contains the text string "E-iceblue".
			For i As Integer = 0 To presentation.Slides.Count - 1
				Dim j As Integer = 0
				Do While j < presentation.Slides(i).Shapes.Count
					If TypeOf presentation.Slides(i).Shapes(j) Is IAutoShape Then
						Dim shape As IAutoShape = TryCast(presentation.Slides(i).Shapes(j), IAutoShape)
						If shape.TextFrame.Text.Contains("E-iceblue") Then
							presentation.Slides(i).Shapes.Remove(shape)
						End If
					End If
					j += 1
				Loop
			Next i

			'Remove image watermark.
			For i As Integer = 0 To presentation.Slides.Count - 1
				presentation.Slides(i).SlideBackground.Fill.FillType = FillFormatType.None
			Next i

			Dim result As String = "Result-RemoveTextAndImageWatermarks.pptx"

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