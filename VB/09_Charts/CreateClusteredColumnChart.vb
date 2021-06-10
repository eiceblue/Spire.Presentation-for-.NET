Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections

Namespace CreateClusteredColumnChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT file
			Dim presentation As New Presentation()

			'Add clustered column chart
			Dim rect1 As New RectangleF(90, 100, 550, 320)
			Dim chart As IChart = presentation.Slides(0).Shapes.AppendChart(ChartType.ColumnClustered, rect1, False)

			'Chart title
			chart.ChartTitle.TextProperties.Text = "Clustered Column Chart"
			chart.ChartTitle.TextProperties.IsCentered = True
			chart.ChartTitle.Height = 30
			chart.HasTitle = True

			'Data for series
			Dim Series1() As Double = { 7.7, 8.9, 1.0, 2.4 }
			Dim Series2() As Double = { 15.2, 5.3, 6.7, 8 }

			'Set series text
			chart.ChartData(0, 1).Text = "Series1"
			chart.ChartData(0, 2).Text = "Series2"

			'Set category text
			chart.ChartData(1, 0).Text = "Category 1"
			chart.ChartData(2, 0).Text = "Category 2"
			chart.ChartData(3, 0).Text = "Category 3"
			chart.ChartData(4, 0).Text = "Category 4"

			'Fill data for chart
			For i As Int32 = 0 To Series1.Length - 1
				chart.ChartData(i + 1, 1).Value = Series1(i)
				chart.ChartData(i + 1, 2).Value = Series2(i)

			Next i

			'Set series label
			chart.Series.SeriesLabel = chart.ChartData("B1", "C1")
			'Set category label
			chart.Categories.CategoryLabels = chart.ChartData("A2", "A5")

			'Set values for series
			chart.Series(0).Values = chart.ChartData("B2", "B5")
			chart.Series(1).Values = chart.ChartData("C2", "C5")

			Dim result As String = "CreateClusteredColumnChart_result.pptx"
			'Save the document
			presentation.SaveToFile(result, FileFormat.Pptx2010)

			'Launch the result file
			PPTDocViewer(result)
		End Sub

		Private Sub PPTDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace