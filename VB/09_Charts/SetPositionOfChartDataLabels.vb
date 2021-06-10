Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace SetPositionOfChartDataLabels
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_2.pptx")

			'Get the chart.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Add data label to chart and set its id.
			Dim label1 As ChartDataLabel = chart.Series(0).DataLabels.Add()
			label1.ID = 0

			'Set the default position of data label. This position is relative to the data markers.
			'label1.Position = ChartDataLabelPosition.OutsideEnd;

			'Set custom position of data label. This position is relative to the default position.
			label1.X = 0.1f
			label1.Y = -0.1f

			'Set label value visible
			label1.LabelValueVisible = True

			'Set legend key invisible
			label1.LegendKeyVisible = False

			'Set category name invisible
			label1.CategoryNameVisible = False

			'Set series name invisible
			label1.SeriesNameVisible = False

			'Set Percentage invisible
			label1.PercentageVisible = False

			'Set border style and fill style of data label
			label1.Line.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			label1.Line.SolidFillColor.Color = Color.Blue
			label1.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			label1.Fill.SolidColor.Color = Color.Orange

			Dim result As String = "Result-SetPositionOfChartDataLabels.pptx"

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