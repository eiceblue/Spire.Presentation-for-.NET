Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections
Imports System.IO


Namespace SetTickMarksInterval
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim ppt As New Presentation()
			Dim inputFile As String = "..\..\..\..\..\..\Data\SetTickMarksInterval.pptx"
			ppt.LoadFromFile(inputFile)
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)
			Dim chartAxis As IChartAxis = chart.PrimaryCategoryAxis
			chartAxis.TickMarkSpacing = 2
			'Save the document
			Dim outputFile As String = "SetTickMarksInterval_out.pptx"
			ppt.SaveToFile(outputFile, FileFormat.Pptx2013)

			'Launch the PPT file
			FileViewer(outputFile)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
