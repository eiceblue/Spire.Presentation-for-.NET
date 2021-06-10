Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace SetTextFontForLegendAndAxis
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPonit document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_2.pptx")

			'Get the chart.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Set the font for the text on Chart Legend area.
			chart.ChartLegend.TextProperties.Paragraphs(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Green
			chart.ChartLegend.TextProperties.Paragraphs(0).DefaultCharacterProperties.LatinFont = New TextFont("Arial Unicode MS")

			'Set the font for the text on Chart Axis area.
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs(0).DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs(0).DefaultCharacterProperties.FontHeight = 10
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs(0).DefaultCharacterProperties.LatinFont = New TextFont("Arial Unicode MS")

			Dim result As String = "Result-SetTextFontOfChartLegendAndChartAxis.pptx"

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