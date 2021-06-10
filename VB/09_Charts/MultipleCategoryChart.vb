Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections

Namespace MultipleCategoryChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT file
			Dim presentation As New Presentation()

			'Add line markers chart
			Dim rect1 As New RectangleF(90, 100, 550, 320)
			Dim chart As IChart = presentation.Slides(0).Shapes.AppendChart(ChartType.ColumnClustered, rect1, False)

			'Chart title
			chart.ChartTitle.TextProperties.Text = "Muli-Category"
			chart.ChartTitle.TextProperties.IsCentered = True
			chart.ChartTitle.Height = 30
			chart.HasTitle = True


			'Data for series
			Dim Series1() As Double = { 7.7, 8.9, 7, 6,7, 8 }

			'Set series text
			chart.ChartData(0, 2).Text = "Series1"

			'Set category text
			chart.ChartData(1, 0).Text = "Grp 1"
			chart.ChartData(3, 0).Text = "Grp 2"
			chart.ChartData(5, 0).Text = "Grp 3"

			chart.ChartData(1, 1).Text = "A"
			chart.ChartData(2, 1).Text = "B"
			chart.ChartData(3, 1).Text = "C"
			chart.ChartData(4, 1).Text = "D"
			chart.ChartData(5, 1).Text = "E"
			chart.ChartData(6, 1).Text = "F"


			'Fill data for chart
			For i As Integer = 0 To Series1.Length - 1
				chart.ChartData(i + 1, 2).Value = Series1(i)

			Next i

			'Set series label
			chart.Series.SeriesLabel = chart.ChartData("C1", "C1")
			'Set category label
			chart.Categories.CategoryLabels = chart.ChartData("A2", "B7")

			'Set values for series
			chart.Series(0).Values = chart.ChartData("C2", "C7")

			'Set if the category axis has multiple levels
			chart.PrimaryCategoryAxis.HasMultiLvlLbl = True
			'Merge same label
			chart.PrimaryCategoryAxis.IsMergeSameLabel = True

			Dim result As String = "MultipleCategoryChart_result.pptx"
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