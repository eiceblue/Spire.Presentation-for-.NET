Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace CreateCylinder3DClusteredChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect2 As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect2)
			presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Insert chart
			Dim rect As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 200, 85, 400, 400)
			Dim chart As IChart = presentation.Slides(0).Shapes.AppendChart(Spire.Presentation.Charts.ChartType.Cylinder3DClustered, rect)

			'Add chart Title
			chart.ChartTitle.TextProperties.Text = "Report"
			chart.ChartTitle.TextProperties.IsCentered = True
			chart.ChartTitle.Height = 30
			chart.HasTitle = True

			'Load data from XML file to datatable
			Dim dataTable As DataTable = LoadData()

			'Load data from datatable to chart
			InitChartData(chart, dataTable)
			chart.Series.SeriesLabel = chart.ChartData("B1", "D1")
			chart.Categories.CategoryLabels = chart.ChartData("A2", "A7")
			chart.Series(0).Values = chart.ChartData("B2", "B7")
			chart.Series(0).Fill.FillType = FillFormatType.Solid
			chart.Series(0).Fill.SolidColor.KnownColor = KnownColors.Brown
			chart.Series(1).Values = chart.ChartData("C2", "C7")
			chart.Series(1).Fill.FillType = FillFormatType.Solid
			chart.Series(1).Fill.SolidColor.KnownColor = KnownColors.Green
			chart.Series(2).Values = chart.ChartData("D2", "D7")
			chart.Series(2).Fill.FillType = FillFormatType.Solid
			chart.Series(2).Fill.SolidColor.KnownColor = KnownColors.Orange

			'Set the 3D rotation
			chart.RotationThreeD.XDegree = 10
			chart.RotationThreeD.YDegree = 10

			'Save the document
			presentation.SaveToFile("CreateCylinder3DClusteredChart_result.pptx", FileFormat.Pptx2010)
			Process.Start("CreateCylinder3DClusteredChart_result.pptx")

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