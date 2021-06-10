Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections
Imports Spire.Presentation.Drawing

Namespace AddShadowEffectForDataLabel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_3.pptx")

			'Get the chart.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Add a data label to the first chart series.
			Dim dataLabels As ChartDataLabelCollection = chart.Series(0).DataLabels
			Dim Label As ChartDataLabel = dataLabels.Add()
			Label.LabelValueVisible = True

			'Add outer shadow effect to the data label.
			Label.Effect.OuterShadowEffect = New OuterShadowEffect()

			'Set shadow color.
			Label.Effect.OuterShadowEffect.ColorFormat.Color = Color.Yellow

			'Set blur.
			Label.Effect.OuterShadowEffect.BlurRadius = 5

			'Set distance.
			Label.Effect.OuterShadowEffect.Distance = 10

			'Set angle.
			Label.Effect.OuterShadowEffect.Direction = 90f

			Dim result As String = "Result-AddShadowEffectToChartDataLabels.pptx"

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