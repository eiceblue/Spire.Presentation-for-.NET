Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace CreatMapChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()

			'Insert a Map chart to the first slide 
			Dim chart As IChart = ppt.Slides(0).Shapes.AppendChart(ChartType.Map, New RectangleF(50, 50, 450, 450), False)
			chart.ChartData(0, 1).Text = "series"

			'Define some data.
			Dim countries() As String = { "China", "Russia", "France", "Mexico", "United States", "India", "Australia" }
			For i As Integer = 0 To countries.Length - 1
				chart.ChartData(i + 1, 0).Text = countries(i)
			Next i
			Dim values() As Integer = { 32, 20, 23, 17, 18, 6, 11 }
			For i As Integer = 0 To values.Length - 1
				chart.ChartData(i + 1, 1).NumberValue = values(i)
			Next i
			chart.Series.SeriesLabel = chart.ChartData(0, 1, 0, 1)
			chart.Categories.CategoryLabels = chart.ChartData(1, 0, 7, 0)
			chart.Series(0).Values = chart.ChartData(1, 1, 7, 1)
			Dim result As String = "Result-CreateMapChart.pptx"

			'Save to file.
			ppt.SaveToFile(result, FileFormat.Pptx2013)

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