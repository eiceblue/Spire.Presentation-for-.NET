Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace CreateWaterFallChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim ppt As New Presentation()

			'Create a WaterFall chart to the first slide
			Dim chart As IChart = ppt.Slides(0).Shapes.AppendChart(ChartType.WaterFall, New RectangleF(50, 50, 500, 400), False)

			'Set series text
			chart.ChartData(0, 1).Text = "Series 1"

			'Set category text
			Dim categories() As String = { "Category 1", "Category 2", "Category 3", "Category 4", "Category 5", "Category 6", "Category 7" }
			For i As Integer = 0 To categories.Length - 1
				chart.ChartData(i + 1, 0).Text = categories(i)
			Next i

			'Fill data for chart
			Dim values() As Double = { 100, 20, 50, -40, 130, -60, 70 }
			For i As Integer = 0 To values.Length - 1
				chart.ChartData(i + 1, 1).NumberValue = values(i)
			Next i

			'Set series labels
			chart.Series.SeriesLabel = chart.ChartData(0, 1, 0, 1)

			'Set categories labels 
			chart.Categories.CategoryLabels = chart.ChartData(1, 0, categories.Length, 0)

			'Assign data to series values
			chart.Series(0).Values = chart.ChartData(1, 1, values.Length, 1)

			'Operate the third datapoint of first series
			Dim chartDataPoint As New ChartDataPoint(chart.Series(0))
			chartDataPoint.Index = 2
			chartDataPoint.SetAsTotal = True
			chart.Series(0).DataPoints.Add(chartDataPoint)

			'Operate the sixth datapoint of first series
			Dim chartDataPoint2 As New ChartDataPoint(chart.Series(0))
			chartDataPoint2.Index = 5
			chartDataPoint2.SetAsTotal = True
			chart.Series(0).DataPoints.Add(chartDataPoint2)
			chart.Series(0).ShowConnectorLines = True
			chart.Series(0).DataLabels.LabelValueVisible = True

			chart.ChartLegend.Position = ChartLegendPositionType.Right
			chart.ChartTitle.TextProperties.Text = "WaterFall"

			'Save the document
			Dim outputFile As String = "WaterFallChartResult.pptx"
			ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
			ppt.Dispose()

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
