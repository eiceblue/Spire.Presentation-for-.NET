Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace SetProjectionOfMapChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			Dim inputFile As String = "..\..\..\..\..\..\Data\SetProjectionOfMapChart.pptx"

			' Create Presentation object and load the file
			Dim ppt As New Presentation()
			ppt.LoadFromFile(inputFile)

			' Get the chart
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(9), IChart)

			' Get the type of projection
			Dim type As ProjectionType = chart.Series(0).ProjectionType

			' Change the tpye of projection
			chart.Series(0).ProjectionType = ProjectionType.Robinson

			' Save to file
			ppt.SaveToFile("SetProjectionOfMapChart_output2.pptx", FileFormat.Pptx2013)

			'Dispose
			ppt.Dispose()

			'System.Diagnostics.Process.Start("SetProjectionOfMapChart_output.pptx");

			Me.Close()
		End Sub

	End Class
End Namespace