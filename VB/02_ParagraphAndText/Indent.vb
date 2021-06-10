Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation.Collections
Imports Spire.Presentation

Namespace Indent
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Indent.pptx")

			Dim shape As IAutoShape = CType(presentation.Slides(0).Shapes(0), IAutoShape)
			Dim paras As ParagraphCollection = shape.TextFrame.Paragraphs

			'Set the paragraph style for first paragraph
			paras(0).Indent = 20
			paras(0).LeftMargin = 10
			paras(0).SpaceAfter = 10

			'Set the paragraph style of the third paragraph 
			paras(2).Indent = -100
			paras(2).LeftMargin = 40
			paras(2).SpaceBefore = 0
			paras(2).SpaceAfter = 0

			'Save the document
			presentation.SaveToFile("Indent_result.pptx", FileFormat.Pptx2010)
			Process.Start("Indent_result.pptx")
		End Sub
	End Class
End Namespace