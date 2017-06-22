Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace Axis
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\BarChart.pptx")

			'get the chart
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'add a secondary axis to display the value of Series 3
			chart.Series(2).UseSecondAxis = True

			'Set the grid line of secondary axis as invisible
			chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None

			'set bounds of axis value. Before we assign values, we must set IsAutoMax and IsAutoMin as false, otherwise MS PowerPoint will automatically set the values.
			chart.PrimaryValueAxis.IsAutoMax = False
			chart.PrimaryValueAxis.IsAutoMin = False
			chart.SecondaryValueAxis.IsAutoMax = False
			chart.SecondaryValueAxis.IsAutoMax = False

			chart.PrimaryValueAxis.MinValue = 0f
			chart.PrimaryValueAxis.MaxValue = 5.0f
			chart.SecondaryValueAxis.MinValue = 0f
			chart.SecondaryValueAxis.MaxValue = 1.0f

			'set axis line format
			chart.PrimaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid
			chart.SecondaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid
			chart.PrimaryValueAxis.MinorGridLines.Width = 0.1f
			chart.SecondaryValueAxis.MinorGridLines.Width = 0.1f
			chart.PrimaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.LightGray
			chart.SecondaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.LightGray
			chart.PrimaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash
			chart.SecondaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash

			chart.PrimaryValueAxis.MajorGridTextLines.Width = 0.3f
			chart.PrimaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.LightSkyBlue
			chart.SecondaryValueAxis.MajorGridTextLines.Width = 0.3f
			chart.SecondaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.LightSkyBlue

			ppt.SaveToFile("secondAxis.pptx", FileFormat.Pptx2010)
			Process.Start("secondAxis.pptx")
		End Sub
	End Class
End Namespace
