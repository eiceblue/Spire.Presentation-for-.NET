Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ResetPositionOfPlaceholder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_7.pptx")

			'Get the first slide from the sample document.
			Dim slide As ISlide = presentation.Slides(0)

			For Each shapeToMove As IShape In slide.Shapes
				'Reset the position of the slide number to the left.
				If shapeToMove.Name.Contains("Slide Number Placeholder") Then
					shapeToMove.Left = 0

				ElseIf shapeToMove.Name.Contains("Date Placeholder") Then
					'Reset the position of the date time to the center.
					shapeToMove.Left = presentation.SlideSize.Size.Width \ 2

					'Reset the date time display style.
					TryCast(shapeToMove, IAutoShape).TextFrame.TextRange.Paragraph.Text = Date.Now.ToString("dd.MM.yyyy")
					TryCast(shapeToMove, IAutoShape).TextFrame.IsCentered = True
				End If
			Next shapeToMove

			Dim result As String = "Result-ResetPositionOfDateTimeAndSlideNumber.pptx"

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