Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation.Charts
Imports System.Text.RegularExpressions

Namespace ReplaceTextWithRegex
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create Presentation
			Dim presentation As New Presentation()

			'Load file
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SomePresentation.pptx")

			'Regex for all words
			Dim regex As New Regex("\d+.\d+|\w+")

			'New string value
			Dim newvalue As String = "This is the test!"

			'Loop and replace
			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					shape.ReplaceTextWithRegex(regex, newvalue)
				Next shape
			Next slide
			'Save the file
			Dim result As String = "ReplaceTextWithRegex_result.pptx"
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