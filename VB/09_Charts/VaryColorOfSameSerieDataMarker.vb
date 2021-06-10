Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace VaryColorOfSameSerieDataMarker
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\VaryColorsOfSameSeriesDataMarkers.pptx")

			'Get the chart from the presentation.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Create a ChartDataPoint object and specify the index.
			Dim dataPoint As New ChartDataPoint(chart.Series(0))
			dataPoint.Index = 0

			'Set the fill color of the data marker.
			dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
			dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Red

			'Set the line color of the data marker.
			dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
			dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.Red

			'Add the data point to the point collection of a series.
			chart.Series(0).DataPoints.Add(dataPoint)

			dataPoint = New ChartDataPoint(chart.Series(0))
			dataPoint.Index = 1
			dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
			dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Black
			dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
			dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.Black
			chart.Series(0).DataPoints.Add(dataPoint)

			dataPoint = New ChartDataPoint(chart.Series(0))
			dataPoint.Index = 2
			dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
			dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Blue
			dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
			dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.Blue
			chart.Series(0).DataPoints.Add(dataPoint)

			Dim result As String = "Result-VaryColorsOfSameSeriesDataMarkers.pptx"

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