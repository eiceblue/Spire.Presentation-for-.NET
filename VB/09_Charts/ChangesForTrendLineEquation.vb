Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation.Charts

Namespace ChangesForTrendLineEquation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create Presentation
			Dim presentation As New Presentation()

			'Load ppt file
			presentation.LoadFromFile("..\..\..\..\..\..\Data\TrendlineEquation.pptx")

			'Get chart on the first slide
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Get the first trendline 
			Dim trendline As ITrendlines = TryCast(chart.Series(0).TrendLines(0), ITrendlines)

			'Change font size for trendline Equation text
			For Each para As TextParagraph In trendline.TrendLineLabel.TextFrameProperties.Paragraphs
				para.DefaultCharacterProperties.FontHeight = 20
				For Each range As Spire.Presentation.TextRange In para.TextRanges
					range.FontHeight = 20
				Next range
			Next para

			'Change position for trendline Equation
			trendline.TrendLineLabel.OffsetX = -0.1f
			trendline.TrendLineLabel.OffsetY = -0.05f

			'Save the file
			Dim result As String = "ChangesForTrendLineEquation_result.pptx"
			presentation.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013)

			'Launching the result file.
			Viewer(result)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace