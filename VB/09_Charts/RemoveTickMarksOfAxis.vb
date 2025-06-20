Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace RemoveTickMarksOfAxis
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

			'Get the chart that need to be adjusted the number format and remove the tick marks.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Set percentage number format for the axis value of chart.
			chart.PrimaryValueAxis.NumberFormat = "0#\%"

			'Remove the tick marks for value axis and category axis.
			chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkNone
			chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkNone
			chart.PrimaryCategoryAxis.MajorTickMark = TickMarkType.TickMarkNone
			chart.PrimaryCategoryAxis.MinorTickMark = TickMarkType.TickMarkNone

			Dim result As String = "Result-SetNumberFormatAndRemoveTickMarksOfChart.pptx"

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