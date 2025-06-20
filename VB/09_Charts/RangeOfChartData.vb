Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace RangeOfChartData
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()

			'Load PPT file 
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ChartSample2.pptx")

			'Create a StringBuilder object
			Dim sb As New StringBuilder()

			'Get chart on the first slide
			Dim chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)
			If chart IsNot Nothing Then
				Dim lastRow As Integer = chart.ChartData.LastRowIndex
				Dim lastCol As Integer = chart.ChartData.LastColIndex
				sb.AppendLine("lastRowIndex: " & lastRow & vbCrLf & "lastColIndex: " & lastCol)
			End If

			'Save to txt file
			Dim result As String = "output.txt"
			File.WriteAllText(result, sb.ToString())

			Process.Start(result)
		End Sub
	End Class
End Namespace