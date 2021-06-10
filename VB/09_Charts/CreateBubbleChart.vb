Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports System.ComponentModel
Imports System.Text

Namespace CreateBubbleChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT file.
			Dim presentation As New Presentation()

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect2 As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect2)
			presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Add bubble chart
			Dim rect1 As New RectangleF(90, 100, 550, 320)
			Dim chart As IChart = presentation.Slides(0).Shapes.AppendChart(ChartType.Bubble, rect1, False)

			'Chart title
			chart.ChartTitle.TextProperties.Text = "Bubble Chart"
			chart.ChartTitle.TextProperties.IsCentered = True
			chart.ChartTitle.Height = 30
			chart.HasTitle = True

			'Attach the data to chart
			Dim xdata() As Double = { 7.7, 8.9, 1.0, 2.4 }
			Dim ydata() As Double = { 15.2, 5.3, 6.7, 8 }
			Dim size() As Double = { 1.1, 2.4, 3.7, 4.8 }

			chart.ChartData(0, 0).Text = "X-Value"
			chart.ChartData(0, 1).Text = "Y-Value"
			chart.ChartData(0, 2).Text = "Size"

			For i As Int32 = 0 To xdata.Length - 1
				chart.ChartData(i + 1, 0).Value = xdata(i)
				chart.ChartData(i + 1, 1).Value = ydata(i)
				chart.ChartData(i + 1, 2).Value = size(i)
			Next i

			'Set series label
			chart.Series.SeriesLabel = chart.ChartData("B1", "B1")

			chart.Series(0).XValues = chart.ChartData("A2", "A5")
			chart.Series(0).YValues = chart.ChartData("B2", "B5")
			chart.Series(0).Bubbles.Add(chart.ChartData("C2"))
			chart.Series(0).Bubbles.Add(chart.ChartData("C3"))
			chart.Series(0).Bubbles.Add(chart.ChartData("C4"))
			chart.Series(0).Bubbles.Add(chart.ChartData("C5"))

			presentation.SaveToFile("BubbleChart_result.pptx", FileFormat.Pptx2010)
			Process.Start("BubbleChart_result.pptx")
		End Sub
	End Class
End Namespace
