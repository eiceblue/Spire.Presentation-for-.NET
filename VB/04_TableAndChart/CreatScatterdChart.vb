Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports System.ComponentModel
Imports System.Text

Namespace CreatScatterdChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'creat a presentation
			Dim pres As New Presentation()
			'insert a chart and set chart title and chart type
			Dim rect1 As New RectangleF(40, 40, 550, 320)
			Dim chart As IChart = pres.Slides(0).Shapes.AppendChart(ChartType.ScatterMarkers, rect1, False)
			chart.ChartTitle.TextProperties.Text = "ScatterMarker Chart"
			chart.ChartTitle.TextProperties.IsCentered = True
			chart.ChartTitle.Height = 30
			chart.HasTitle = True

			'set chart data
			Dim xdata() As Double = { 2.7, 8.9, 10.0, 12.4 }
			Dim ydata() As Double = { 3.2, 15.3, 6.7, 8 }

			chart.ChartData(0, 0).Text = "X-Value"
			chart.ChartData(0, 1).Text = "Y-Value"

			For i As Int32 = 0 To xdata.Length - 1
				chart.ChartData(i + 1, 0).Value = xdata(i)
				chart.ChartData(i + 1, 1).Value = ydata(i)
			Next i

			'set the series label
			chart.Series.SeriesLabel = chart.ChartData("B1", "B1")

			'assign data to X axis, Y axis and Bubbles
			chart.Series(0).XValues = chart.ChartData("A2", "A5")
			chart.Series(0).YValues = chart.ChartData("B2", "B5")


			pres.SaveToFile("ScatterMarkerChart.pptx", FileFormat.Pptx2010)
			Process.Start("ScatterMarkerChart.pptx")
		End Sub
	End Class
End Namespace
