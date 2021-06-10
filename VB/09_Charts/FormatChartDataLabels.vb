Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections
Imports System.ComponentModel
Imports System.Text

Namespace FormatChartDataLabels
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file.
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\FormatChartDataLabels.pptx")

			'Get the chart
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Get the chart series
			Dim sers As ChartSeriesFormatCollection = chart.Series

			'Initialize four instances of series label and set parameters of each label
			Dim cd1 As ChartDataLabel = sers(0).DataLabels.Add()
			cd1.PercentageVisible = True
			cd1.TextFrame.Text = "Custom Datalabel1"
			cd1.TextFrame.TextRange.FontHeight = 12
			cd1.TextFrame.TextRange.LatinFont = New TextFont("Lucida Sans Unicode")
			cd1.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			cd1.TextFrame.TextRange.Fill.SolidColor.Color= Color.Green

			Dim cd2 As ChartDataLabel = sers(0).DataLabels.Add()
			cd2.Position = ChartDataLabelPosition.InsideEnd
			cd2.PercentageVisible = True
			cd2.TextFrame.Text = "Custom Datalabel2"
			cd2.TextFrame.TextRange.FontHeight = 10
			cd2.TextFrame.TextRange.LatinFont = New TextFont("Arial")
			cd2.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			cd2.TextFrame.TextRange.Fill.SolidColor.Color = Color.OrangeRed

			Dim cd3 As ChartDataLabel = sers(0).DataLabels.Add()
			cd3.Position = ChartDataLabelPosition.Center
			cd3.PercentageVisible = True
			cd3.TextFrame.Text = "Custom Datalabel3"
			cd3.TextFrame.TextRange.FontHeight = 14
			cd3.TextFrame.TextRange.LatinFont = New TextFont("Calibri")
			cd3.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			cd3.TextFrame.TextRange.Fill.SolidColor.Color = Color.Blue

			Dim cd4 As ChartDataLabel = sers(0).DataLabels.Add()
			cd4.Position = ChartDataLabelPosition.InsideBase
			cd4.PercentageVisible = True
			cd4.TextFrame.Text = "Custom Datalabel4"
			cd4.TextFrame.TextRange.FontHeight = 12
			cd4.TextFrame.TextRange.LatinFont = New TextFont("Lucida Sans Unicode")
			cd4.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			cd4.TextFrame.TextRange.Fill.SolidColor.Color = Color.OliveDrab

			'Save and launch the file 
			ppt.SaveToFile("FormatDataLable_result.pptx", FileFormat.Pptx2010)
			Process.Start("FormatDataLable_result.pptx")
		End Sub
	End Class
End Namespace
