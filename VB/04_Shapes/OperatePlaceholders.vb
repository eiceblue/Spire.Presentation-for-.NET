Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation
Imports Spire.Presentation.Diagrams
Imports Spire.Presentation.Charts

Namespace OperatePlaceholders
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\OperatePlaceholders.pptx")

			'Operate placeholders
			For j As Integer = 0 To presentation.Slides.Count - 1
				Dim slide As ISlide = CType(presentation.Slides(j), ISlide)

				For i As Integer = 0 To slide.Shapes.Count - 1
					Dim shape As Shape = CType(slide.Shapes(i), Shape)
					Select Case shape.Placeholder.Type
						Case PlaceholderType.Media
							shape.InsertVideo("..\..\..\..\..\..\Data\Video.mp4")

						Case PlaceholderType.Picture
							shape.InsertPicture("..\..\..\..\..\..\Data\E-iceblueLogo.png")

						Case PlaceholderType.Chart
							shape.InsertChart(ChartType.ColumnClustered)

						Case PlaceholderType.Table
							shape.InsertTable(3,2)

						Case PlaceholderType.Diagram
							shape.InsertSmartArt(SmartArtLayoutType.BasicBlockList)
					End Select
				Next i
			Next j

			Dim result As String="OperatePlaceholders_result.pptx"
			'Save the document
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			PPTDocumentViewer(result)
		End Sub
		Private Sub PPTDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try

		End Sub
	End Class
End Namespace