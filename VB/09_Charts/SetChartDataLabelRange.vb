Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace SetChartDataLabelRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Add a ColumnStacked chart
			Dim chart As IChart = presentation.Slides(0).Shapes.AppendChart(ChartType.ColumnStacked, New RectangleF(100, 100, 500, 400))

			'Set data for the chart
			Dim cellRange As CellRange = chart.ChartData("F1")
			cellRange.Text = "labelA"
			cellRange = chart.ChartData("F2")
			cellRange.Text = "labelB"
			cellRange = chart.ChartData("F3")
			cellRange.Text = "labelC"
			cellRange = chart.ChartData("F4")
			cellRange.Text = "labelD"

			'Set data label ranges
			chart.Series(0).DataLabelRanges = chart.ChartData("F1", "F4")

			'Add data label
			Dim dataLabel1 As ChartDataLabel = chart.Series(0).DataLabels.Add()
			dataLabel1.ID = 0
			'Show the value
			dataLabel1.LabelValueVisible = False
			'Show the label string
			dataLabel1.ShowDataLabelsRange = True

			Dim result As String = "Result-SetChartDataLabelRange.pptx"
			'Save to file
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file
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