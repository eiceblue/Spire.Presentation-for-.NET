Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace HideOrShowASeriesOfChart
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

			'Get the first slide.
			Dim slide As ISlide = presentation.Slides(0)

			'Get the first chart.
			Dim chart As IChart = TryCast(slide.Shapes(0), IChart)

			'Hide the first series of the chart.
			chart.Series(0).IsHidden = True

			'Show the first series of the chart.
			'chart.Series[0].IsHidden = false;

			Dim result As String = "Result-HideOrShowASeriesOfChart.pptx"

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