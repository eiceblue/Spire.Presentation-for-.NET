Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetBordersForNewlyTables
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Set the table width and height for each table cell.
			Dim tableWidth() As Double = { 100, 100, 100, 100, 100 }
			Dim tableHeight() As Double = { 20, 20 }

			'Traverse all the border type of the table.
			For Each item As TableBorderType In System.Enum.GetValues(GetType(TableBorderType))
			  'Add a table to the presentation slide with the setting width and height
				Dim itable As ITable = presentation.Slides.Append().Shapes.AppendTable(100, 100, tableWidth, tableHeight)

				'Add some text to the table cell.
				itable.TableRows(0)(0).TextFrame.Text = "Row"
				itable.TableRows(1)(0).TextFrame.Text = "Column"

				'Set the border type, border width and the border color for the table.
				itable.SetTableBorder(item, 1.5, Color.Red)
			Next item

			Dim result As String = "Result-SetBordersForNewlyTables.pptx"

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