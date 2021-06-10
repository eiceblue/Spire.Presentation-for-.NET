Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace SetAxisType
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SetAxisType.pptx")

			'Get the chart
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(1), IChart)

			chart.PrimaryCategoryAxis.AxisType = Spire.Presentation.Charts.AxisType.DateAxis
			chart.PrimaryCategoryAxis.MajorUnitScale = ChartBaseUnitType.Months

			Dim result As String = "SetAxisType_result.pptx"

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