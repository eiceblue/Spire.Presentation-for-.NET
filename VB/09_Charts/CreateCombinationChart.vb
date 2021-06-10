Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace CreateCombinationChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a presentation instance
			Dim presentation As New Presentation()

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect2 As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect2)
			presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Insert a column clustered chart
			Dim rect As New RectangleF(100, 100, 550, 320)
			Dim chart As IChart = presentation.Slides(0).Shapes.AppendChart(ChartType.ColumnClustered, rect)

			'Set chart title
			chart.ChartTitle.TextProperties.Text = "Monthly Sales Report"
			chart.ChartTitle.TextProperties.IsCentered = True
			chart.ChartTitle.Height = 30
			chart.HasTitle = True

			'Create a datatable
			Dim dataTable As New DataTable()
			dataTable.Columns.Add(New DataColumn("Month", Type.GetType("System.String")))
			dataTable.Columns.Add(New DataColumn("Sales", Type.GetType("System.Int32")))
			dataTable.Columns.Add(New DataColumn("Growth rate", Type.GetType("System.Decimal")))
			dataTable.Rows.Add("January", 200, 0.6)
			dataTable.Rows.Add("February", 250, 0.8)
			dataTable.Rows.Add("March", 300, 0.6)
			dataTable.Rows.Add("April", 150, 0.2)
			dataTable.Rows.Add("May", 200, 0.5)
			dataTable.Rows.Add("June", 400, 0.9)

			'Import data from datatable to chart data
			For c As Integer = 0 To dataTable.Columns.Count - 1
				chart.ChartData(0, c).Text = dataTable.Columns(c).Caption
			Next c
			For r As Integer = 0 To dataTable.Rows.Count - 1
				Dim datas() As Object = dataTable.Rows(r).ItemArray
				For c As Integer = 0 To datas.Length - 1
					chart.ChartData(r + 1, c).Value = datas(c)

				Next c
			Next r

			'Set series labels
			chart.Series.SeriesLabel = chart.ChartData("B1", "C1")

			'Set categories labels    
			chart.Categories.CategoryLabels = chart.ChartData("A2", "A7")

			'Assign data to series values
			chart.Series(0).Values = chart.ChartData("B2", "B7")
			chart.Series(1).Values = chart.ChartData("C2", "C7")

			'Change the chart type of serie 2 to line with markers
			chart.Series(1).Type = ChartType.LineMarkers

			'Plot data of series 2 on the secondary axis
			chart.Series(1).UseSecondAxis = True

			'Set the number format as percentage 
			chart.SecondaryValueAxis.NumberFormat = "0%"

			'Hide gridlinkes of secondary axis
			chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None

			'Set overlap
			chart.OverLap = -50

			'Set gapwidth
			chart.GapWidth = 200

			'Save to file
			presentation.SaveToFile("CombinationChart_result.pptx", FileFormat.Pptx2010)
			Process.Start("CombinationChart_result.pptx")
		End Sub
	End Class
End Namespace
