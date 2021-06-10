Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace ChangeTextFontInChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPTX file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ChangeTextFontInChart.pptx")

			'Get the chart
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Change the font of title
			chart.ChartTitle.TextProperties.Paragraphs(0).DefaultCharacterProperties.LatinFont = New TextFont("Lucida Sans Unicode")
			chart.ChartTitle.TextProperties.Paragraphs(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Blue
			chart.ChartTitle.TextProperties.Paragraphs(0).DefaultCharacterProperties.FontHeight = 30

			'Change the font of legend
			chart.ChartLegend.TextProperties.Paragraphs(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.DarkGreen
			chart.ChartLegend.TextProperties.Paragraphs(0).DefaultCharacterProperties.LatinFont = New TextFont("Lucida Sans Unicode")

			'Change the font of series
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs(0).DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs(0).DefaultCharacterProperties.FontHeight = 10
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs(0).DefaultCharacterProperties.LatinFont = New TextFont("Lucida Sans Unicode")

			ppt.SaveToFile("ChangeTextFontInChart_result.pptx", FileFormat.Pptx2010)
			Process.Start("ChangeTextFontInChart_result.pptx")
		End Sub
	End Class
End Namespace
