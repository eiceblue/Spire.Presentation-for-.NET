Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace Alignment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Alignment.pptx")

			'Get the related shape and set the text alignment
			Dim shape As IAutoShape = CType(presentation.Slides(0).Shapes(1), IAutoShape)
			shape.TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Left
			shape.TextFrame.Paragraphs(1).Alignment = TextAlignmentType.Center
			shape.TextFrame.Paragraphs(2).Alignment = TextAlignmentType.Right
			shape.TextFrame.Paragraphs(3).Alignment = TextAlignmentType.Justify
			shape.TextFrame.Paragraphs(4).Alignment = TextAlignmentType.None

			'Save the document
			presentation.SaveToFile("alignment_result.pptx", FileFormat.Pptx2010)
			Process.Start("alignment_result.pptx")
		End Sub
	End Class
End Namespace