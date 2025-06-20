Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace CreateTreeMapChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim ppt As New Presentation()

			'Create a TreeMap chart to the first slide
			Dim chart As IChart = ppt.Slides(0).Shapes.AppendChart(ChartType.TreeMap, New RectangleF(50, 50, 500, 400), False)

			'Set series text
			chart.ChartData(0, 3).Text = "Series 1"

			'Set category text
			Dim categories(,) As String = {{"Branch 1","Stem 1","Leaf 1"},{"Branch 1","Stem 1","Leaf 2"},{"Branch 1","Stem 1", "Leaf 3"}, {"Branch 1","Stem 2","Leaf 4"},{"Branch 1","Stem 2","Leaf 5"},{"Branch 1","Stem 2","Leaf 6"},{"Branch 1","Stem 2","Leaf 7"}, {"Branch 2","Stem 3","Leaf 8"},{"Branch 2","Stem 3","Leaf 9"},{"Branch 2","Stem 4","Leaf 10"},{"Branch 2","Stem 4","Leaf 11"}, {"Branch 2","Stem 5","Leaf 12"},{"Branch 3","Stem 5","Leaf 13"},{"Branch 3","Stem 6","Leaf 14"},{"Branch 3","Stem 6","Leaf 15"}}
			For i As Integer = 0 To 14
				For j As Integer = 0 To 2
					chart.ChartData(i + 1, j).Text = categories(i, j)
				Next j
			Next i

			'Fill data for chart
			Dim values() As Double = { 17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51 }
			For i As Integer = 0 To values.Length - 1
				chart.ChartData(i + 1, 3).NumberValue = values(i)
			Next i

			'Set series labels
			chart.Series.SeriesLabel = chart.ChartData(0, 3, 0, 3)

			'Set categories labels 
			chart.Categories.CategoryLabels = chart.ChartData(1, 0, values.Length, 2)

			'Assign data to series values
			chart.Series(0).Values = chart.ChartData(1, 3, values.Length, 3)

			chart.Series(0).DataLabels.CategoryNameVisible = True
			chart.Series(0).TreeMapLabelOption = TreeMapLabelOption.Banner
			chart.ChartTitle.TextProperties.Text = "TreeMap"
			chart.HasLegend = True
			chart.ChartLegend.Position = ChartLegendPositionType.Top

			'Save the document
			Dim outputFile As String = "TreeMapChartResult.pptx"
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
