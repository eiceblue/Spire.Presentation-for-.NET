Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation.Charts

Namespace SetChartBorderstyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create Presentation
			Dim presentation As New Presentation()

			'Load ppt file
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ChartSample2.pptx")

			'Get chart on the first slide
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Set border style
			chart.Line.FillFormat.FillType = FillFormatType.Solid
			chart.Line.FillFormat.SolidFillColor.Color = Color.Red
			chart.BorderRoundedCorners = True

			'Save the file
			Dim result As String = "SetChartBorderstyle_result.pptx"
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