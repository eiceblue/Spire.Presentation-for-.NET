Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing


Namespace CreateParetoChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim ppt As New Presentation()

			'Create a Pareto chart in first slide
			Dim chart As IChart = ppt.Slides(0).Shapes.AppendChart(ChartType.Pareto, New RectangleF(50, 50, 500, 400), False)

			'Set series text
			chart.ChartData(0, 1).Text = "Series 1"

			'Set category text
			Dim categories() As String = { "Category 1", "Category 2", "Category 4", "Category 3", "Category 4", "Category 2", "Category 1", "Category 1", "Category 3", "Category 2", "Category 4", "Category 2", "Category 3", "Category 1", "Category 3", "Category 2", "Category 4", "Category 1", "Category 1", "Category 3", "Category 2", "Category 4", "Category 1", "Category 1", "Category 3", "Category 2", "Category 4", "Category 1"}
			For i As Integer = 0 To categories.Length - 1
				chart.ChartData(i + 1, 0).Text = categories(i)
			Next i

			'Fill data for chart
			Dim values() As Double = { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 }
			For i As Integer = 0 To values.Length - 1
				chart.ChartData(i + 1, 1).NumberValue = values(i)
			Next i

			chart.Series.SeriesLabel = chart.ChartData(0, 1, 0, 1)
			chart.Categories.CategoryLabels = chart.ChartData(1, 0, categories.Length, 0)
			chart.Series(0).Values = chart.ChartData(1, 1, values.Length, 1)
			chart.PrimaryCategoryAxis.IsBinningByCategory = True
			chart.Series(1).Line.FillFormat.FillType = FillFormatType.Solid
			chart.Series(1).Line.FillFormat.SolidFillColor.Color = Color.Red
			chart.ChartTitle.TextProperties.Text = "Pareto"
			chart.HasLegend = True
			chart.ChartLegend.Position = ChartLegendPositionType.Bottom

			'Save the document
			Dim outputFile As String = "ParetoChartResult.pptx"
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
