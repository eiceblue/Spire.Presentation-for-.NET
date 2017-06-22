Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections
Imports System.ComponentModel
Imports System.Text

Namespace FormatChartDataLabels
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create PPT document and load file.
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\PieChart.pptx")

			'get the chart
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'get the chart series
			Dim sers As ChartSeriesFormatCollection = chart.Series

			'set the chart legend position to Right
			chart.ChartLegend.Position = ChartLegendPositionType.Right

			'initialize four instances of series label and set parameters of each label
			Dim cd1 As ChartDataLabel = sers(0).DataLabels.Add()
			cd1.Position = ChartDataLabelPosition.Center
			cd1.PercentageVisible = True

			Dim cd2 As ChartDataLabel = sers(0).DataLabels.Add()
			cd2.PercentageVisible = True
			cd2.Position = ChartDataLabelPosition.Center

			Dim cd3 As ChartDataLabel = sers(0).DataLabels.Add()
			cd3.PercentageVisible = True
			cd3.Position = ChartDataLabelPosition.Center

			Dim cd4 As ChartDataLabel = sers(0).DataLabels.Add()
			cd4.PercentageVisible = True
			cd4.Position = ChartDataLabelPosition.Center

			ppt.SaveToFile("FormatDataLable.pptx", FileFormat.Pptx2010)
			Process.Start("FormatDataLable.pptx")
		End Sub
	End Class
End Namespace
