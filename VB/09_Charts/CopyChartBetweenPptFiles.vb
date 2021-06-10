Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace CopyChartBetweenPptFiles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation1 As New Presentation()

			'Load the file from disk.
			presentation1.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_2.pptx")

			'Get the chart that is going to be copied.
			Dim chart As IChart = TryCast(presentation1.Slides(0).Shapes(0), IChart)

			'Load the second PowerPoint document.
			Dim presentation2 As New Presentation()
			presentation2.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_1.pptx")

			'Copy chart from the first document to the second document.
			presentation2.Slides.Append()
			presentation2.Slides(1).Shapes.CreateChart(chart, New RectangleF(100, 100, 500, 300), -1)

			Dim result As String = "Result-CopyChartBetweenPptFiles.pptx"

			'Save to file.
			presentation2.SaveToFile(result, FileFormat.Pptx2013)

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