Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace GroupTwoLevelAxisLabels
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\GroupTwoLevelAxisLabels.pptx")

			'Get the chart.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Get the category axis from the chart.
			Dim chartAxis As IChartAxis = chart.PrimaryCategoryAxis

			'Group the axis labels that have the same first-level label.
			If chartAxis.HasMultiLvlLbl Then
				chartAxis.IsMergeSameLabel = True
			End If

			Dim result As String = "Result-GroupTwoLevelAxisLabels.pptx"

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