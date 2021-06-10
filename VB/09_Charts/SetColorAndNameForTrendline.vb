Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace SetColorAndNameForTrendline
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a ppt document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\SetColorAndNameForTrendline.pptx")

			'Find the first chart in the first Slide
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Find the first trendline in the chart
			Dim trendline As ITrendlines = TryCast(chart.Series(0).TrendLines(0), ITrendlines)

			'Set name for trendline
			trendline.Name = "trendlineName"

			'Set color for trendline
			trendline.Line.FillType = FillFormatType.Solid
			trendline.Line.SolidFillColor.Color = Color.Red

			'Save the document
			ppt.SaveToFile("SetColorAndNameForTrendline_result.pptx", FileFormat.Pptx2010)

			'Launch the file
			OutputViewer("SetColorAndNameForTrendline_result.pptx")
		End Sub

		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace