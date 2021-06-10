Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts

Namespace RemoveChartFromPptSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPonit document
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_3.pptx")

			'Get the first slide from the document.
			Dim slide As ISlide = presentation.Slides(0)

			'Remove chart from the slide.
			Dim i As Integer = 0
			Do While i < slide.Shapes.Count
				Dim shape As IShape = TryCast(slide.Shapes(i), IShape)
				If TypeOf shape Is IChart Then
					slide.Shapes.Remove(shape)
				End If
				i += 1
			Loop

			Dim result As String = "Result-RemoveChartFromPptSlide.pptx"

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