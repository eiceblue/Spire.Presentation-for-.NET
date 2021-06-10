Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace AddRowToTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_1.pptx")

			'Get the table within the PowerPoint document.
			Dim table As ITable = TryCast(presentation.Slides(0).Shapes(0), ITable)

			'Get the second row.
			Dim row As TableRow = table.TableRows(1)

			'Clone the row and add it to the end of table.
			table.TableRows.Append(row)
			Dim rowCount As Integer = table.TableRows.Count

			'Get the last row.
			Dim lastRow As TableRow = table.TableRows(rowCount - 1)

			'Set new data of the first cell of last row.
			lastRow(0).TextFrame.Text = " The first added cell"

			'Set new data of the second cell of last row.
			lastRow(1).TextFrame.Text = " The second added cell"

			Dim result As String = "Result-AddRowToTable.pptx"

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