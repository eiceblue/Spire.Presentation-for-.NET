Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace SetTickMarkLabelsOnCategoryAxis
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPonit document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_3.pptx")

			'Get the chart from the PowerPoint slide.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Rotate tick labels.
			chart.PrimaryCategoryAxis.TextRotationAngle = 45

			'Specify interval between labels.
			chart.PrimaryCategoryAxis.IsAutomaticTickLabelSpacing = False
			chart.PrimaryCategoryAxis.TickLabelSpacing = 2

			'Change position.
			chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionHigh

			Dim result As String = "Result-SetTickMarkLabelsOnCategoryAxis.pptx"

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