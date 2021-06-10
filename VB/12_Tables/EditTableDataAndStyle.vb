Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace EditTableDataAndStyle
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

			'Store the data used in replacement in string [].
			Dim str() As String = { "Germany", "Berlin", "Europe", "0152458", "20860000" }

			Dim table As ITable = Nothing

			'Get the table in PowerPoint document.
			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ITable Then
					table = CType(shape, ITable)

					'Change the style of table.
					table.StylePreset = TableStylePreset.LightStyle1Accent2

					For i As Integer = 0 To table.ColumnsList.Count - 1
						'Replace the data in cell.
						table(i, 2).TextFrame.Text = str(i)

						'Set the highlightcolor.
						table(i, 2).TextFrame.TextRange.HighlightColor.Color = Color.BlueViolet
					Next i
				End If
			Next shape

			Dim result As String = "Result-EditTableDataAndStyle.pptx"

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