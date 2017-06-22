Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace RemoveRow
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\table.pptx")

			'get the table in PPT document
			Dim table As ITable = Nothing
			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ITable Then
					table = CType(shape, ITable)

					'remove the second column
					table.ColumnsList.RemoveAt(1, False)

					'remove the second row
					table.TableRows.RemoveAt(1, False)
				End If
			Next shape
			'save the document
			presentation.SaveToFile("RemoveRow.pptx", FileFormat.Pptx2010)
			Process.Start("RemoveRow.pptx")
		End Sub

		Private Sub lblDescription_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lblDescription.Click

		End Sub

		Private Sub pbLogo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles pbLogo.Click

		End Sub
	End Class
End Namespace
