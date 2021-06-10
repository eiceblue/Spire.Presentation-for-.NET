Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections

Namespace ModifyChartCategoryAxis
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

			'Modify the major unit
			Chart.PrimaryCategoryAxis.IsAutoMajor = False
			Chart.PrimaryCategoryAxis.MajorUnit = 1
			Chart.PrimaryCategoryAxis.MajorUnitScale = ChartBaseUnitType.Months

			Dim result As String = "ModifyChartCategoryAxis_result.pptx"
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