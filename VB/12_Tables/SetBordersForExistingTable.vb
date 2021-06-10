Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetBordersForExistingTable
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

			'Get the table from the first slide of the sample document.
			Dim slide As ISlide = presentation.Slides(0)
			Dim table As ITable = TryCast(slide.Shapes(0), ITable)

			'Set the border type as Inside and the border color as blue.
			table.SetTableBorder(TableBorderType.Inside, 1, Color.Blue)

			Dim result As String = "Result-SetBordersForExistingTable.pptx"

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