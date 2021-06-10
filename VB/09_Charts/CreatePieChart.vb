Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace CreatePieChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Insert a Pie chart to the first slide and set the chart title.
			Dim rect1 As New RectangleF(40, 100, 550, 320)
			Dim chart As IChart = presentation.Slides(0).Shapes.AppendChart(ChartType.Pie, rect1, False)
			chart.ChartTitle.TextProperties.Text = "Sales by Quarter"
			chart.ChartTitle.TextProperties.IsCentered = True
			chart.ChartTitle.Height = 30
			chart.HasTitle = True

			'Define some data.
			Dim quarters() As String = { "1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr" }
			Dim sales() As Integer = { 210, 320, 180, 500 }

			'Append data to ChartData, which represents a data table where the chart data is stored.
			chart.ChartData(0, 0).Text = "Quarters"
			chart.ChartData(0, 1).Text = "Sales"
			For i As Integer = 0 To quarters.Length - 1
				chart.ChartData(i + 1, 0).Value = quarters(i)
				chart.ChartData(i + 1, 1).Value = sales(i)
			Next i

			'Set category labels, series label and series data.
			chart.Series.SeriesLabel = chart.ChartData("B1", "B1")
			chart.Categories.CategoryLabels = chart.ChartData("A2", "A5")
			chart.Series(0).Values = chart.ChartData("B2", "B5")

			'Add data points to series and fill each data point with different color.
			For i As Integer = 0 To chart.Series(0).Values.Count - 1
				Dim cdp As New ChartDataPoint(chart.Series(0))
				cdp.Index = i
				chart.Series(0).DataPoints.Add(cdp)

			Next i
			chart.Series(0).DataPoints(0).Fill.FillType = FillFormatType.Solid
			chart.Series(0).DataPoints(0).Fill.SolidColor.Color = Color.RosyBrown
			chart.Series(0).DataPoints(1).Fill.FillType = FillFormatType.Solid
			chart.Series(0).DataPoints(1).Fill.SolidColor.Color = Color.LightBlue
			chart.Series(0).DataPoints(2).Fill.FillType = FillFormatType.Solid
			chart.Series(0).DataPoints(2).Fill.SolidColor.Color = Color.LightPink
			chart.Series(0).DataPoints(3).Fill.FillType = FillFormatType.Solid
			chart.Series(0).DataPoints(3).Fill.SolidColor.Color = Color.MediumPurple

			'Set the data labels to display label value and percentage value.
			chart.Series(0).DataLabels.LabelValueVisible = True
			chart.Series(0).DataLabels.PercentValueVisible = True

			Dim result As String = "Result-CreatePieChart.pptx"

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