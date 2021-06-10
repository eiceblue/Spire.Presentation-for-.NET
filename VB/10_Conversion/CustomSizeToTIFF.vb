Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO

Namespace CustomSizeToTIFF
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the original PPT document from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Indent.pptx")

			'Get the first slide
			Dim slide As ISlide= presentation.Slides(0)

			'Create a new PPT document
			Dim newPresentation As New Presentation()

			'Remove the default slide 
			newPresentation.Slides.RemoveAt(0)

			'Define a new size
			Dim size As New SizeF(200F, 200F)

			'Set PPT slide size
			newPresentation.SlideSize.Size = size

			'Insert the slide of original PPT
			newPresentation.Slides.Insert(0, slide)

			'String for output file 
			Dim result As String = "Output1.tiff"

			'Save the second slide to PDF
			newPresentation.SaveToFile(result, Spire.Presentation.FileFormat.Tiff)


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