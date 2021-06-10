Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace SetRotationForDataLabel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\SetRotationForDataLabel.pptx")

			'Get chart on the first slide
			Dim Chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Set the rotation angle for the datalabels of first serie
			For i As Integer = 0 To Chart.Series(0).Values.Count - 1
				Dim datalabel As ChartDataLabel = Chart.Series(0).DataLabels.Add()
				datalabel.ID = i
				datalabel.RotationAngle = 45
			Next i

			Dim result As String = "SetRotationForDataLabel_out.pptx"

			'Save the document
			ppt.SaveToFile(result, FileFormat.Pptx2013)

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