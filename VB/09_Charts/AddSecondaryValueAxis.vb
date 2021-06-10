Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace AddSecondaryValueAxis
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_2.pptx")

			'Get the chart from the PowerPoint file.
			Dim chart As IChart = TryCast(presentation.Slides(0).Shapes(0), IChart)

			'Add a secondary axis to display the value of Series 3.
			chart.Series(2).UseSecondAxis = True

			'Set the grid line of secondary axis as invisible.
			chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None

			Dim result As String = "Result-AddSecondaryValueAxisToChart.pptx"

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