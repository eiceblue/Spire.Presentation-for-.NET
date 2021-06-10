Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SplitSpecificTableCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_1.pptx")

			'Get the first slide.
			Dim slide As ISlide = presentation.Slides(0)

			'Get the table.
			Dim table As ITable = TryCast(slide.Shapes(0), ITable)

			'Split cell [1, 2] into 3 rows and 2 columns.
			table(1, 2).Split(3, 2)

			Dim result As String = "Result-SplitSpecificTableCell.pptx"

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