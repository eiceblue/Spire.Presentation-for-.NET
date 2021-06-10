Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace CloneRowAndColumn
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim presentation As New Presentation()
			' Access first slide
			Dim sld As ISlide = presentation.Slides(0)

			' Define columns with widths and rows with heights
			Dim widths() As Double = { 110, 110, 110 }
			Dim heights() As Double = { 50, 30, 30, 30, 30 }

			' Add table shape to slide
			Dim table As ITable = presentation.Slides(0).Shapes.AppendTable(presentation.SlideSize.Size.Width \ 2 - 275, 90, widths, heights)

			' Add text to the row 1 cell 1
			table(0, 0).TextFrame.Text = "Row 1 Cell 1"

			' Add text to the row 1 cell 2
			table(1, 0).TextFrame.Text = "Row 1 Cell 2"

			' Clone row 1 at end of table
			table.TableRows.Append(table.TableRows(0))

			' Add text to the row 2 cell 1
			table(0, 1).TextFrame.Text = "Row 2 Cell 1"

			' Add text to the row 2 cell 2
			table(1, 1).TextFrame.Text = "Row 2 Cell 2"

			' Clone row 2 as the 4th row of table
			table.TableRows.Insert(3, table.TableRows(1))

			'Clone column 1 at end of table
			table.ColumnsList.Add(table.ColumnsList(0))

			'Clone the 2nd column at 4th column index
			table.ColumnsList.Insert(3, table.ColumnsList(1))

			Dim result As String = "CloneRowAndColumn_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2010)

			Viewer(result)
		End Sub

		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace