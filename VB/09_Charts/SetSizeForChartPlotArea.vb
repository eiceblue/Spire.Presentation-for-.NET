Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports System.IO

Namespace SetSizeForChartPlotArea
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ChartSample2.pptx")

			'Get chart on the first slide
			Dim Chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Set width and height for chart plot area
			Chart.PlotArea.Width = 250
			Chart.PlotArea.Height = 300


			Dim result As String = "SetSizeForChartPlotArea_result.pptx"
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

	End Class
End Namespace