Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace RemoveRowsAndColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\RemoveRowsAndColumns.pptx")

			'Get the table in PPT document
			Dim table As ITable = Nothing
			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ITable Then
					table = CType(shape, ITable)

					'Remove the second column
					table.ColumnsList.RemoveAt(1, False)

					'Remove the second row
					table.TableRows.RemoveAt(1, False)
				End If
			Next shape
			'Save and launch the document
			presentation.SaveToFile("RemoveRowsAndColumns_result.pptx", FileFormat.Pptx2010)
			Process.Start("RemoveRowsAndColumns_result.pptx")
		End Sub
	End Class
End Namespace
