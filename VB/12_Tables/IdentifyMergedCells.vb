Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace IdentifyMergedCells
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\MergedCellInTable.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			Dim str As New StringBuilder()
			Dim output As String=""
			For Each shape As IShape In slide.Shapes
				'Verify if it is table
				If TypeOf shape Is ITable Then
					Dim table As ITable = CType(shape, ITable)
					For r As Integer = 0 To table.TableRows.Count - 1
						For c As Integer = 0 To table.ColumnsList.Count - 1
							' Get cell
							Dim currentCell As Cell = table.TableRows(r)(c)
							'Identify if it is merged cell
							If currentCell.RowSpan>1 OrElse currentCell.ColSpan>1 Then
								output =String.Format("Cell {0}:{1} is a part of merged cell with RowSpan={2} and ColSpan={3} starting from Cell {4}:{5}.", r, c, currentCell.RowSpan, currentCell.ColSpan, currentCell.FirstRowIndex, currentCell.FirstColumnIndex)

								str.AppendLine(output)
							End If
						Next c

					Next r
				End If
			Next shape

			Dim result As String = "IdentifyMergedCells_result.txt"
			File.WriteAllText(result, str.ToString())
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