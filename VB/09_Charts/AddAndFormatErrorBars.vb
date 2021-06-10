Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace AddAndFormatErrorBars
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AddAndFormatErrorBars.pptx")

			'Get the column chart on the first slide and set chart title.
			Dim columnChart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			columnChart.ChartTitle.TextProperties.Text = "Vertical Error Bars"

		   'Add Y (Vertical) Error Bars.

			'Get Y error bars of the first chart series.
			Dim errorBarsYFormat1 As IErrorBarsFormat = columnChart.Series(0).ErrorBarsYFormat

			'Set end cap.
			errorBarsYFormat1.ErrorBarNoEndCap = False

			'Specify direction.
			errorBarsYFormat1.ErrorBarSimType = ErrorBarSimpleType.Plus

			'Specify error amount type.
			errorBarsYFormat1.ErrorBarvType = ErrorValueType.StandardError

			'Set value.
			errorBarsYFormat1.ErrorBarVal = 0.3f

			'Set line format.
			errorBarsYFormat1.Line.FillType = FillFormatType.Solid
			errorBarsYFormat1.Line.SolidFillColor.Color = Color.MediumVioletRed
			errorBarsYFormat1.Line.Width = 1

			'Get the bubble chart on the second slide and set chart title.
			Dim bubbleChart As IChart = TryCast(presentation.Slides(1).Shapes(0), IChart)

			bubbleChart.ChartTitle.TextProperties.Text = "Vertical and Horizontal Error Bars"


			 'Add X (Horizontal) and Y (Vertical) Error Bars.
			'Get X error bars of the first chart series.
			Dim errorBarsXFormat As IErrorBarsFormat = bubbleChart.Series(0).ErrorBarsXFormat

			'Set end cap.
			errorBarsXFormat.ErrorBarNoEndCap = False

			'Specify direction.
			errorBarsXFormat.ErrorBarSimType = ErrorBarSimpleType.Both

			'Specify error amount type.
			errorBarsXFormat.ErrorBarvType = ErrorValueType.StandardError

			'Set value.
			errorBarsXFormat.ErrorBarVal = 0.3f

			'Get Y error bars of the first chart series.
			Dim errorBarsYFormat2 As IErrorBarsFormat = bubbleChart.Series(0).ErrorBarsYFormat

			'Set end cap.
			errorBarsYFormat2.ErrorBarNoEndCap = False

			'Specify direction.
			errorBarsYFormat2.ErrorBarSimType = ErrorBarSimpleType.Both

			'Specify error amount type.
			errorBarsYFormat2.ErrorBarvType = ErrorValueType.StandardError

			'Set value.
			errorBarsYFormat2.ErrorBarVal = 0.3f

			Dim result As String = "Result-AddAndFormatErrorBars.pptx"

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