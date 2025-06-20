Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Security.Policy
Imports System.Text

Namespace DetectIsSwitchRowAndColumn
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ChartSample5.pptx")

			'Get the chart
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Detect whether the chart has "SwitchRowAndColumn" setting
			Dim result As Boolean = chart.IsSwitchRowAndColumn()

			MessageBox.Show("'SwitchRowAndColumn' value of the chart is " & result)
		End Sub
	End Class
End Namespace
