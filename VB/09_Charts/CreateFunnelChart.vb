Imports Spire.Presentation
Imports Spire.Presentation.Charts


Namespace CreateFunnelChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim ppt As New Presentation()

			'Create a Funnel chart to the first slide
			Dim chart As IChart = ppt.Slides(0).Shapes.AppendChart(ChartType.Funnel, New RectangleF(50, 50, 550, 400), False)

			'Set series text
			chart.ChartData(0, 1).Text = "Series 1"

			'Set category text
			Dim categories() As String = { "Website Visits", "Download", "Uploads", "Requested price", "Invoice sent", "Finalized" }
			For i As Integer = 0 To categories.Length - 1
				chart.ChartData(i + 1, 0).Text = categories(i)
			Next i

			'Fill data for chart
			Dim values() As Double = { 50000, 47000, 30000, 15000, 9000, 5600 }
			For i As Integer = 0 To values.Length - 1
				chart.ChartData(i + 1, 1).NumberValue = values(i)
			Next i

			'Set series labels
			chart.Series.SeriesLabel = chart.ChartData(0, 1, 0, 1)

			'Set categories labels 
			chart.Categories.CategoryLabels = chart.ChartData(1, 0, categories.Length, 0)

			'Assign data to series values
			chart.Series(0).Values = chart.ChartData(1, 1, values.Length, 1)

			'Set the chart title
			chart.ChartTitle.TextProperties.Text = "Funnel"

			Dim outputFile As String = "FunnelChartResult.pptx"
			'Save the document
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
