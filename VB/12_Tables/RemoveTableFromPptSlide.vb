Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace RemoveTableFromPptSlide
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

			'Get the tables within the PPT document.
			Dim shape_tems As New List(Of IShape)()

			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ITable Then
					'Add new table to table list.
					shape_tems.Add(shape)
				End If
			Next shape

			'Remove all the tables form the first slide.
			For Each shape As IShape In shape_tems
				presentation.Slides(0).Shapes.Remove(shape)
			Next shape

			Dim result As String = "Result-RemoveTableFromPptSlide.pptx"

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