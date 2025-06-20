Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports System.ComponentModel

Namespace SetDistanceFromAxis
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a ppt document
			Dim ppt As New Presentation()

			'Append ColumnClustered chart
			Dim chart As IChart = ppt.Slides(0).Shapes.AppendChart(ChartType.ColumnClustered, New RectangleF(50, 50, 400, 400))

			'Get the PrimaryCategory axis
			Dim chartAxis As IChartAxis = chart.PrimaryCategoryAxis

			'Set "Distance from axis"
			chartAxis.LabelsDistance = 200

			'Save the file
			ppt.SaveToFile("SetDistanceFromAxis.pptx", FileFormat.Pptx2013)

			'Launch and view the resulted PPTX file
			PresentationDocViewer("SetDistanceFromAxis.pptx")
		End Sub
		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
