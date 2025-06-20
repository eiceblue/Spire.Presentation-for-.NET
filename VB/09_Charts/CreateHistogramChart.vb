Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace CreateHistogramChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim ppt As New Presentation()

			'Add a Histogram chart
			Dim chart As IChart = ppt.Slides(0).Shapes.AppendChart(ChartType.Histogram, New RectangleF(50, 50, 500, 400), False)

			'Set series text
			chart.ChartData(0, 0).Text = "Series 1"

			'Fill data for chart
			Dim values() As Double = { 1, 1, 1, 3, 3, 3, 3, 5, 5, 5, 8, 8, 8, 9, 9, 9, 12, 12, 13, 13, 17, 17, 17, 19, 19, 19, 25, 25, 25, 25, 25, 25, 25, 25, 29, 29, 29, 29, 32, 32, 33, 33, 35, 35, 41, 41, 44, 45, 49, 49 }
			For i As Integer = 0 To values.Length - 1
				chart.ChartData(i + 1, 1).NumberValue = values(i)
			Next i

			'Set series label
			chart.Series.SeriesLabel = chart.ChartData(0, 0, 0, 0)

			'Set values for series
			chart.Series(0).Values = chart.ChartData(1, 0, values.Length, 0)

			chart.PrimaryCategoryAxis.NumberOfBins = 7
			chart.PrimaryCategoryAxis.GapWidth = 20
			'Chart title
			chart.ChartTitle.TextProperties.Text = "Histogram"
			chart.ChartLegend.Position = ChartLegendPositionType.Bottom

			Dim outputFile As String = "histogramChartResult.pptx"
			'Save the document
			ppt.SaveToFile(outputFile, FileFormat.Pptx2013)

			'Launch the PPT file
			FileViewer(outputFile)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
