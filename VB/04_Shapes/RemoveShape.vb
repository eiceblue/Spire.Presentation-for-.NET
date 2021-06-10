Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace RemoveShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load doucment from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\FindShapeByAltText.pptx")

			'Loop through slides
			For i As Integer = 0 To presentation.Slides.Count - 1
				Dim slide As ISlide = presentation.Slides(i)
				'Loop through shapes
				Dim j As Integer = 0
				Do While j < slide.Shapes.Count
					Dim shape As IShape = slide.Shapes(j)
					'Find the shapes whose alternative text contain "Shape"
					If shape.AlternativeText.Contains("Shape") Then
						slide.Shapes.Remove(shape)
						j -= 1
					End If
					j += 1
				Loop
			Next i

			'Save the document
			Dim result As String = "RemoveShape_result.pptx"
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