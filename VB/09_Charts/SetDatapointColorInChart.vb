Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace SetDatapointColorInChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\SetDatapointColorInChart.pptx")

			'Get the chart
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Initialize an instances of dataPoint
			Dim cdp1 As New ChartDataPoint(chart.Series(0))

			'Specify the datapoint order
			cdp1.Index = 0

			'Set the color of the datapoint
			cdp1.Fill.FillType = FillFormatType.Solid
			cdp1.Fill.SolidColor.KnownColor = KnownColors.Orange

			'Add the dataPoint to first series
			chart.Series(0).DataPoints.Add(cdp1)

			'Set the color for the other three data points
			Dim cdp2 As New ChartDataPoint(chart.Series(0))
			cdp2.Index = 1
			cdp2.Fill.FillType = FillFormatType.Solid
			cdp2.Fill.SolidColor.KnownColor = KnownColors.Gold
			chart.Series(0).DataPoints.Add(cdp2)

			Dim cdp3 As New ChartDataPoint(chart.Series(0))
			cdp3.Index = 2
			cdp3.Fill.FillType = FillFormatType.Solid
			cdp3.Fill.SolidColor.KnownColor = KnownColors.MediumPurple
			chart.Series(0).DataPoints.Add(cdp3)

			Dim cdp4 As New ChartDataPoint(chart.Series(0))
			cdp4.Index = 1
			cdp4.Fill.FillType = FillFormatType.Solid
			cdp4.Fill.SolidColor.KnownColor = KnownColors.ForestGreen
			chart.Series(0).DataPoints.Add(cdp4)


			ppt.SaveToFile("SetDatapointColor_result.pptx", FileFormat.Pptx2010)
			Process.Start("SetDatapointColor_result.pptx")
		End Sub
	End Class
End Namespace
