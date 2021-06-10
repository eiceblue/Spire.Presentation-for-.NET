Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace AddTrendLineForChartSeries
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_2.pptx")

			'Get the target chart, add trendline for the first data series of the chart and specify the trendline type.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)
			Dim it As ITrendlines = chart.Series(0).AddTrendLine(TrendlinesType.Linear)

			'Set the trendline properties to determine what should be displayed.
			it.displayEquation = False
			it.displayRSquaredValue = False

			Dim result As String = "Result-AddTrendLineForChartSeries.pptx"

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