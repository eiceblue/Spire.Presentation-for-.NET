Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections

Namespace SetPercentageForLabels
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ColumnStacked.pptx")

			'Get the chart on the first slide
			Dim Chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

		   Dim dataPontPercent As Single = 0f

		   For i As Integer = 0 To Chart.Series.Count - 1
			   Dim series As ChartSeriesDataFormat = Chart.Series(i)
			   'Get the total number
			   Dim total As Single = GetTotal(series.Values)
			   For j As Integer = 0 To series.Values.Count - 1
				 'Get the percent
				 dataPontPercent = Single.Parse(series.Values(j).Text) / total * 100
				 'Add datalabels
				 Dim label As ChartDataLabel = series.DataLabels.Add()
				 label.LabelValueVisible = True
				 'Set the percent text for the label
				 label.TextFrame.Paragraphs(0).Text = String.Format("{0:F2} %", dataPontPercent)
				 label.TextFrame.Paragraphs(0).TextRanges(0).FontHeight = 12
			   Next j
		   Next i

		   Dim result As String = "SetPercentageForLabels_result.pptx"
			'Save the document
			ppt.SaveToFile(result, FileFormat.Pptx2010)

			'Launch the result file
			PPTDocViewer(result)

		End Sub

		Private Sub PPTDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Function GetTotal(ByVal ranges As CellRanges) As Single
			Dim total As Single = 0
			For i As Integer = 0 To ranges.Count - 1
				total += Single.Parse(ranges(i).Text)
			Next i

		   Return total
		End Function
	End Class
End Namespace