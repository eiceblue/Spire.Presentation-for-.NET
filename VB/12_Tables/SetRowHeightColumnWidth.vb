Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace SetRowHeightColumnWidth
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Creat a ppt document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\SetRowHeightColumnWidth.pptx")

			'Get the table
			Dim table As ITable = Nothing
			For Each shape As IShape In ppt.Slides(0).Shapes
				If TypeOf shape Is ITable Then
					table = CType(shape, ITable)

					'Set the height for the rows
					table.TableRows(0).Height = 100
					table.TableRows(1).Height = 80
					table.TableRows(2).Height = 60
					table.TableRows(3).Height = 40
					table.TableRows(4).Height = 20

					'Set the column width
					table.ColumnsList(0).Width = 60
					table.ColumnsList(1).Width = 80
					table.ColumnsList(2).Width = 120
					table.ColumnsList(3).Width = 140
					table.ColumnsList(4).Width = 160
				End If
			Next shape
			'Save the file
			ppt.SaveToFile("RowHeightAndColumnWidth_result.pptx", FileFormat.Pptx2010)
			Process.Start("RowHeightAndColumnWidth_result.pptx")
		End Sub
	End Class
End Namespace
