Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace SetChartDataNumberFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\SetChartDataNumberFormat.pptx")

			'Get chart on the first slide
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Set the number format for Axis
			chart.PrimaryValueAxis.NumberFormat = "#,##0.00"

			'Set the DataLabels format for Axis
			chart.Series(0).DataLabels.LabelValueVisible = True
			chart.Series(0).DataLabels.PercentValueVisible = False
			chart.Series(0).DataLabels.NumberFormat = "#,##0.00"
			chart.Series(0).DataLabels.HasDataSource = False

			'Set the number format for ChartData
			For i As Integer = 1 To chart.Series(0).Values.Count
				chart.ChartData(i, 1).NumberFormat = "#,##0.00"
			Next i

			Dim result As String = "SetChartDataNumberFormat_output.pptx"

			'Save the document
			ppt.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the result file
			PPTDocViewer(result)
		End Sub

		Private Sub PPTDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace