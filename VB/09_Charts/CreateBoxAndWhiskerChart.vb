Imports Spire.Presentation
Imports Spire.Presentation.Charts


Namespace CreateBoxAndWhiskerChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a PPT document
			Dim ppt As New Presentation()

			' Insert a BoxAndWhisker chart to the first slide 
			Dim chart As IChart = ppt.Slides(0).Shapes.AppendChart(ChartType.BoxAndWhisker, New RectangleF(50, 50, 500, 400), False)

			' Series labels
			Dim seriesLabel() As String = { "Series 1", "Series 2", "Series 3" }
			For i As Integer = 0 To seriesLabel.Length - 1
				chart.ChartData(0, i + 1).Text = "Series 1"
			Next i

			' Categories
			Dim categories() As String = {"Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 2", "Category 2", "Category 2", "Category 2", "Category 2", "Category 2", "Category 3", "Category 3", "Category 3", "Category 3", "Category 3"}
			For i As Integer = 0 To categories.Length - 1
				chart.ChartData(i + 1, 0).Text = categories(i)
			Next i

			' Values
			Dim values(,) As Double = {{-7,-3,-24},{-10,1,11},{-28,-6,34},{47,2,-21},{35,17,22},{-22,15,19},{17,-11,25}, {-30,18,25},{49,22,56},{37,22,15},{-55,25,31},{14,18,22},{18,-22,36},{-45,25,-17}, {-33,18,22},{18,2,-23},{-33,-22,10},{10,19,22}}
			For i As Integer = 0 To seriesLabel.Length - 1
				For j As Integer = 0 To categories.Length - 1
					chart.ChartData(j + 1, i + 1).NumberValue = values(j, i)
				Next j
			Next i

			chart.Series.SeriesLabel = chart.ChartData(0, 1, 0, seriesLabel.Length)
			chart.Categories.CategoryLabels = chart.ChartData(1, 0, categories.Length, 0)

			chart.Series(0).Values = chart.ChartData(1, 1, categories.Length, 1)
			chart.Series(1).Values = chart.ChartData(1, 2, categories.Length, 2)
			chart.Series(2).Values = chart.ChartData(1, 3, categories.Length, 3)

			chart.Series(0).ShowInnerPoints = False
			chart.Series(0).ShowOutlierPoints = True
			chart.Series(0).ShowMeanMarkers = True
			chart.Series(0).ShowMeanLine = True
			chart.Series(0).QuartileCalculationType = QuartileCalculation.ExclusiveMedian

			chart.Series(1).ShowInnerPoints = False
			chart.Series(1).ShowOutlierPoints = True
			chart.Series(1).ShowMeanMarkers = True
			chart.Series(1).ShowMeanLine = True
			chart.Series(1).QuartileCalculationType = QuartileCalculation.InclusiveMedian

			chart.Series(2).ShowInnerPoints = False
			chart.Series(2).ShowOutlierPoints = True
			chart.Series(2).ShowMeanMarkers = True
			chart.Series(2).ShowMeanLine = True
			chart.Series(2).QuartileCalculationType = QuartileCalculation.ExclusiveMedian

			chart.HasLegend = True
			chart.ChartTitle.TextProperties.Text = "BoxAndWhisker"
			chart.ChartLegend.Position = ChartLegendPositionType.Top

			Dim outputFile As String = "result.pptx"
			'Save to file
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
