Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports System.IO

Namespace GetValuesAndUnitFromAxis
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim sb As New StringBuilder()

			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ChartSample2.pptx")

			'Get chart on the first slide
			Dim Chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Get unit from primary category axis
			Dim MajorUnit As Single = Chart.PrimaryCategoryAxis.MajorUnit
			Dim type As ChartBaseUnitType = Chart.PrimaryCategoryAxis.MajorUnitScale

			sb.Append(MajorUnit.ToString() & vbCrLf)
			sb.Append(type.ToString() & vbCrLf)


			'Get values from primary value axis
			Dim minValue As Single = Chart.PrimaryValueAxis.MinValue
			Dim maxValue As Single = Chart.PrimaryValueAxis.MaxValue

			sb.Append(minValue.ToString() & vbCrLf)
			sb.Append(maxValue.ToString() & vbCrLf)


			Dim result As String = "GetValuesAndUnitFromAxis_result.txt"
			'Save the document
			File.WriteAllText(result, sb.ToString())

			'Launch the result file
			PPTDocViewer(result)

		End Sub

		Private Sub PPTDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try

		End Sub

	End Class
End Namespace