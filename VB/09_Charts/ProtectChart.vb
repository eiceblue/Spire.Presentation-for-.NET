Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace ProtectChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPonit document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_2.pptx")

			'Get the first shape from slide and convert it as IChart.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Set the Boolean value of IChart.IsDataProtect as true.
			chart.IsDataProtect = True

			Dim result As String = "Result-ProtectChart.pptx"

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