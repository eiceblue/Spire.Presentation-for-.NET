Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO

Namespace IndividualSlideToHtml
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT document from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ChangeSlidePosition.pptx")

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'String for output file 
			Dim result As String = "Output.html"

			'Save the first slide to HTML 
			slide.SaveToFile(result, Spire.Presentation.FileFormat.Html)

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