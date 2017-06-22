Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace SetDatapointColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\Chart.pptx")

			'get the chart
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'initialize an instances of dataPoint
			Dim cdp As New ChartDataPoint(chart.Series(0))

			'specific the dataPoint
			cdp.Index = 2

			'fill the dataPoint
			cdp.Fill.FillType = FillFormatType.Solid
			cdp.Fill.SolidColor.KnownColor = KnownColors.Yellow

			'add the dataPoint to first series
			chart.Series(0).DataPoints.Add(cdp)

			ppt.SaveToFile("SetDatapointColor.pptx", FileFormat.Pptx2010)
			Process.Start("SetDatapointColor.pptx")
		End Sub
	End Class
End Namespace
