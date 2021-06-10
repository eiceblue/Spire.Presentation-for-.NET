Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.IO

Namespace TraverseThroughCells
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPonit document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_1.pptx")

			Dim content As New StringBuilder()
			content.AppendLine("The data in cells of this PowerPoint file is: ")

			'Get the table.
			Dim table As ITable = Nothing
			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ITable Then
					table = CType(shape, ITable)

					'Traverse through the cells of table.
					For Each row As TableRow In table.TableRows
						For Each cell As Cell In row
							content.AppendLine(cell.TextFrame.Text)
						Next cell
						content.AppendLine(vbLf)
					Next row
				End If
			Next shape

			Dim result As String = "Result-TraverseThroughCells.txt"

			'Save to file.
			File.WriteAllText(result, content.ToString())

			'Launch the file.
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