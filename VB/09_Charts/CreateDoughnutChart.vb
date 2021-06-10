Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace CreateDoughnutChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a ppt document
			Dim presentation As New Presentation()
			Dim rect As New RectangleF(80, 100, 550, 320)

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect2 As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect2)
			presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Add a Doughnut chart
			Dim chart As IChart = presentation.Slides(0).Shapes.AppendChart(ChartType.Doughnut, rect, False)
			chart.ChartTitle.TextProperties.Text = "Market share by country"
			chart.ChartTitle.TextProperties.IsCentered = True
			chart.ChartTitle.Height = 30

			Dim countries() As String = { "Guba", "Mexico", "France", "German" }
			Dim sales() As Integer = { 1800, 3000, 5100, 6200 }
			chart.ChartData(0, 0).Text = "Countries"
			chart.ChartData(0, 1).Text = "Sales"
			For i As Integer = 0 To countries.Length - 1
				chart.ChartData(i + 1, 0).Value = countries(i)
				chart.ChartData(i + 1, 1).Value = sales(i)
			Next i
			chart.Series.SeriesLabel = chart.ChartData("B1", "B1")
			chart.Categories.CategoryLabels = chart.ChartData("A2", "A5")
			chart.Series(0).Values = chart.ChartData("B2", "B5")

			For i As Integer = 0 To chart.Series(0).Values.Count - 1
				Dim cdp As New ChartDataPoint(chart.Series(0))
				cdp.Index = i
				chart.Series(0).DataPoints.Add(cdp)
			Next i
			'Set the series color
			chart.Series(0).DataPoints(0).Fill.FillType = FillFormatType.Solid
			chart.Series(0).DataPoints(0).Fill.SolidColor.Color = Color.LightBlue
			chart.Series(0).DataPoints(1).Fill.FillType = FillFormatType.Solid
			chart.Series(0).DataPoints(1).Fill.SolidColor.Color = Color.MediumPurple
			chart.Series(0).DataPoints(2).Fill.FillType = FillFormatType.Solid
			chart.Series(0).DataPoints(2).Fill.SolidColor.Color = Color.DarkGray
			chart.Series(0).DataPoints(3).Fill.FillType = FillFormatType.Solid
			chart.Series(0).DataPoints(3).Fill.SolidColor.Color = Color.DarkOrange

			chart.Series(0).DataLabels.LabelValueVisible = True
			chart.Series(0).DataLabels.PercentValueVisible = True
			chart.Series(0).DoughnutHoleSize = 60

			presentation.SaveToFile("DoughnutChart_result.pptx", FileFormat.Pptx2013)
			Process.Start("DoughnutChart_result.pptx")

		End Sub

		'Function to load data from XML file to DataTable
		Private Function LoadData() As DataTable
			Dim ds As New DataSet()
			ds.ReadXmlSchema("..\..\..\..\..\..\Data\data-schema.xml")
			ds.ReadXml("..\..\..\..\..\..\Data\data.xml")

			Return ds.Tables(0)
		End Function

		'Function to load data from DataTable to IChart
		Private Sub InitChartData(ByVal chart As IChart, ByVal dataTable As DataTable)
			For c As Integer = 0 To dataTable.Columns.Count - 1
				chart.ChartData(0, c).Text = dataTable.Columns(c).Caption
			Next c

			For r As Integer = 0 To dataTable.Rows.Count - 1
				Dim data() As Object = dataTable.Rows(r).ItemArray
				For c As Integer = 0 To data.Length - 1
					chart.ChartData(r + 1, c).Value = data(c)
				Next c
			Next r
		End Sub
	End Class
End Namespace
