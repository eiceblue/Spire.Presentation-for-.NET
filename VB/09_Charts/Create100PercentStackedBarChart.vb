Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace Create100PercentStackedBarChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Add a "Bar100PercentStacked" chart to the first slide.
			presentation.SlideSize.Type = SlideSizeType.Screen16x9
			Dim slidesize As SizeF = presentation.SlideSize.Size

			Dim slide = presentation.Slides(0)

			'Append a chart.
			Dim rect As New RectangleF(20, 20, slidesize.Width - 40, slidesize.Height - 40)
			Dim chart As IChart = slide.Shapes.AppendChart(Spire.Presentation.Charts.ChartType.Bar100PercentStacked, rect)

			'Write data to the chart data.
			Dim columnlabels() As String = { "Series 1", "Series 2", "Series 3" }

			'Insert the column labels.
			For c As Int32 = 0 To columnlabels.Length - 1
				chart.ChartData(0, c + 1).Text = columnlabels(c)
			Next c

			Dim rowlabels() As String = { "Category 1", "Category 2", "Category 3" }

			'Insert the row labels.
			For r As Int32 = 0 To rowlabels.Length - 1
				chart.ChartData(r + 1, 0).Text = rowlabels(r)
			Next r

			Dim values(,) As Double = { { 20.83233, 10.34323, -10.354667 }, { 10.23456, -12.23456, 23.34456 }, { 12.34345, -23.34343, -13.23232 } }

			'Insert the values.
			Dim value As Double = 0.0
			For r As Int32 = 0 To rowlabels.Length - 1
				For c As Int32 = 0 To columnlabels.Length - 1
					value = Math.Round(values(r, c), 2)
					chart.ChartData(r + 1, c + 1).Value = value
				Next c
			Next r

			chart.Series.SeriesLabel = chart.ChartData(0, 1, 0, columnlabels.Length)
			chart.Categories.CategoryLabels = chart.ChartData(1, 0, rowlabels.Length, 0)

			'Set the position of category axis.
			chart.PrimaryCategoryAxis.Position = AxisPositionType.Left
			chart.SecondaryCategoryAxis.Position = AxisPositionType.Left
			chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionLow

			'Set the data, font and format for the series of each column.
			For c As Int32 = 0 To columnlabels.Length - 1
				chart.Series(c).Values = chart.ChartData(1, c + 1, rowlabels.Length, c + 1)
				chart.Series(c).Fill.FillType = FillFormatType.Solid
				chart.Series(c).InvertIfNegative = False

				For r As Int32 = 0 To rowlabels.Length - 1
					Dim label = chart.Series(c).DataLabels.Add()
					label.LabelValueVisible = True
					chart.Series(c).DataLabels(r).HasDataSource = False
					chart.Series(c).DataLabels(r).NumberFormat = "0#\%"
					chart.Series(c).DataLabels.TextProperties.Paragraphs(0).DefaultCharacterProperties.FontHeight = 12
				Next r
			Next c

			'Set the color of the Series.
			chart.Series(0).Fill.SolidColor.Color = Color.YellowGreen
			chart.Series(1).Fill.SolidColor.Color = Color.Red
			chart.Series(2).Fill.SolidColor.Color = Color.Green

			Dim font As New TextFont("Tw Cen MT")

			'Set the font and size for chartlegend.
			For k As Integer = 0 To chart.ChartLegend.EntryTextProperties.Length - 1
				chart.ChartLegend.EntryTextProperties(k).LatinFont = font
				chart.ChartLegend.EntryTextProperties(k).FontHeight = 20
			Next k

			Dim result As String = "Result-Create100PercentStackedBarChart.pptx"

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