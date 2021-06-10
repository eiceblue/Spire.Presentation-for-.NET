Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace HideShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\FindShapeByAltText.pptx")

			'Loop through slides
			For Each slide As ISlide In presentation.Slides
				'Loop through shapes in the slide
				For Each shape As IShape In slide.Shapes
					'Find the shape whose alternative text is Shape1
					If shape.AlternativeText.CompareTo("Shape1") = 0 Then
						'Hide the shape
						shape.IsHidden = True
					End If
				Next shape
			Next slide

			'Save the document
			Dim result As String = "HideShape_result.pptx"
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