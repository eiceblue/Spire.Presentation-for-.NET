Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace SetSizeAndStyleForMarker
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SetSizeAndStyleForMarker.pptx")

			'Get the chart from the presentation.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			For i As Integer = 0 To chart.Series(0).Values.Count - 1
				'Create a ChartDataPoint object and specify the index.
				Dim dataPoint As New ChartDataPoint(chart.Series(0))
				dataPoint.Index = i

				'Set the fill color of the data marker.
				dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
				dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Yellow

				'Set the line color of the data marker.
				dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
				dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.YellowGreen

				'Set the size of the data marker.
				dataPoint.MarkerSize = 20

				'Set the style of the data marker
				dataPoint.MarkerStyle = ChartMarkerType.Diamond
				chart.Series(0).DataPoints.Add(dataPoint)
			Next i

			Dim result As String = "SetSizeAndStyleForMarker_out.pptx"

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