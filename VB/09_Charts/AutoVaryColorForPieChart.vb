Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace AutoVaryColorForPieChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT file
			Dim ppt As New Presentation()

			Dim rect1 As New RectangleF(40, 100, 550, 320)

			'Add a pie chart
			Dim chart As IChart = ppt.Slides(0).Shapes.AppendChart(ChartType.Pie, rect1, False)
			chart.ChartTitle.TextProperties.Text = "Sales by Quarter"
			chart.ChartTitle.TextProperties.IsCentered = True
			chart.ChartTitle.Height = 30
			chart.HasTitle = True

			'Attach the data to chart
			Dim quarters() As String = { "1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr" }
			Dim sales() As Integer = { 210, 320, 180, 500 }
			chart.ChartData(0, 0).Text = "Quarters"
			chart.ChartData(0, 1).Text = "Sales"
			For i As Integer = 0 To quarters.Length - 1
				chart.ChartData(i + 1, 0).Value = quarters(i)
				chart.ChartData(i + 1, 1).Value = sales(i)
			Next i

			chart.Series.SeriesLabel = chart.ChartData("B1", "B1")
			chart.Categories.CategoryLabels = chart.ChartData("A2", "A5")
			chart.Series(0).Values = chart.ChartData("B2", "B5")


			'Set whether auto vary color, default value is true
			chart.Series(0).IsVaryColor = False

			chart.Series(0).Distance = 15

			Dim result As String = "AutoVaryColorForPieChart_result.pptx"
			'Save the document
			ppt.SaveToFile(result, FileFormat.Pptx2010)

			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace