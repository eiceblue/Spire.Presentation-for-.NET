Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections

Namespace AddCustomErrorBars
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.pptx")

			'Get the bubble chart on the first slide
			Dim bubbleChart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Get X error bars of the first chart series
			Dim errorBarsXFormat As IErrorBarsFormat = bubbleChart.Series(0).ErrorBarsXFormat
			'Specify error amount type as custom error bars
			errorBarsXFormat.ErrorBarvType = ErrorValueType.CustomErrorBars
			'Set the minus and plus value of the X error bars
			errorBarsXFormat.MinusVal = 0.5f
			errorBarsXFormat.PlusVal = 0.5f

			'Get Y error bars of the first chart series
			Dim errorBarsYFormat As IErrorBarsFormat = bubbleChart.Series(0).ErrorBarsYFormat
			'Specify error amount type as custom error bars
			errorBarsYFormat.ErrorBarvType = ErrorValueType.CustomErrorBars
			'Set the minus and plus value of the Y error bars
			errorBarsYFormat.MinusVal = 1f
			errorBarsYFormat.PlusVal = 1f

			Dim result As String = "AddCustomErrorBars_result.pptx"
			'Save the document
			ppt.SaveToFile(result, FileFormat.Pptx2010)

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