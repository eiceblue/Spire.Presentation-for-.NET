Imports Spire.Presentation

Namespace AdjustColumnByTextWidth
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

			'Adjust the first column width of table by text width.
			table.ColumnsList(0).AdjustColumnByTextWidth()



			'Save to file.
			Dim result As String = "output.pptx"
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