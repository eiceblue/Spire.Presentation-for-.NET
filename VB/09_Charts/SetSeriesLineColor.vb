Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Text

Namespace SetSeriesLineColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\SeriesLinesColor.pptx")

			'Get the first chart
			Dim shape As IShape = ppt.Slides(0).Shapes(0)
			If TypeOf shape Is IChart Then
				Dim chart As IChart = CType(shape, IChart)
				Dim seriesLine As TextLineFormat = chart.SeriesLine
				seriesLine.FillType = FillFormatType.Solid

				'Set the color of seriesLine
				seriesLine.FillFormat.SolidFillColor.Color = Color.Red
			End If

			'Save the PPT document
			Dim result As String = "SeriesLinesColor_output.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			PresentationDocViewer(result)
		End Sub
		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
