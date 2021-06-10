Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace SetAlignmentInTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SetAlignmentInTable.pptx")

			Dim table As ITable = Nothing
			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ITable Then
					table = CType(shape, ITable)

					'Horizontal Alignment
					'Set the horizontal alignment for the cells in first column 
					table(0, 1).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Left
					table(0, 2).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Center
					table(0, 3).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Right
					table(0, 4).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Justify

					'Vertical Alignment
					'Set the vertical alignment for the cells in second column 
					table(1, 1).TextAnchorType = TextAnchorType.Top
					table(1, 2).TextAnchorType = TextAnchorType.Center
					table(1, 3).TextAnchorType = TextAnchorType.Bottom
					table(1, 4).TextAnchorType = TextAnchorType.None

					'Both orientaions
					'Set the both horizontal and vertical alignment for the cells in the third column 
					table(2, 1).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Left
					table(2, 1).TextAnchorType = TextAnchorType.Top

					table(2, 2).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Right
					table(2, 2).TextAnchorType = TextAnchorType.Center

					table(2, 3).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Justify
					table(2, 3).TextAnchorType = TextAnchorType.Bottom

					table(2, 4).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Center
					table(2, 4).TextAnchorType = TextAnchorType.Top
				End If
			Next shape

			'Save the document
			presentation.SaveToFile("SetAlignmentInTable_result.pptx", FileFormat.Pptx2010)
			Process.Start("SetAlignmentInTable_result.pptx")
		End Sub
	End Class
End Namespace
