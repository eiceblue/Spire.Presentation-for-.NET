Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace CopyChartWithinOnePPT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_2.pptx")

			'Get the chart that is going to be copied.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Copy the chart from the first slide to the specified location of the second slide within the same document.
			Dim slide1 As ISlide = presentation.Slides.Append()
			slide1.Shapes.CreateChart(chart, New RectangleF(100, 100, 500, 300), 0)

			Dim result As String = "Result-CopyChartWithinAPptFile.pptx"

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