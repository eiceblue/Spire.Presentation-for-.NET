Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace SetAndGetAlternativeText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ShapeTemplate.pptx")

			'Get the first slide
			Dim slide As ISlide = ppt.Slides(0)

			'Set the alternative text (title and description)
			slide.Shapes(0).AlternativeTitle = "Rectangle"
			slide.Shapes(0).AlternativeText = "This is a Rectangle"

			'Get the alternative text (title and description)
			Dim alternativeText As String = Nothing
			Dim title As String = slide.Shapes(0).AlternativeTitle
			alternativeText &= "Title: " & title & vbCrLf
			Dim description As String = slide.Shapes(0).AlternativeText
			alternativeText &= "Description: " & description

			'Save the document
			Dim result As String = "SetAlternativeText.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			PresentationDocViewer(result)

			'Save the alternative text to Text file
			result = "GetAlternativeText.txt"
			File.WriteAllText(result, alternativeText)
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