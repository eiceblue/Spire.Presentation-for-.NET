Imports Spire.Presentation.Collections
Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ChangeTextStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ChangeTextStyle.pptx")

			Dim shape As IAutoShape = CType(presentation.Slides(0).Shapes(0), IAutoShape)
			Dim paras As ParagraphCollection = shape.TextFrame.Paragraphs

			'Set the style for the text content in the first paragraph
			For Each tr As TextRange In paras(0).TextRanges
				tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
				tr.Fill.SolidColor.Color = Color.ForestGreen
				tr.LatinFont = New TextFont("Lucida Sans Unicode")
				tr.FontHeight = 14
			Next tr

			'Set the style for the text content in the third paragraph
			For Each tr As TextRange In paras(2).TextRanges
				tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
				tr.Fill.SolidColor.Color = Color.CornflowerBlue
				tr.LatinFont = New TextFont("Calibri")
				tr.FontHeight = 16
				tr.TextUnderlineType = TextUnderlineType.Dashed
			Next tr

			'Save the document
			presentation.SaveToFile("ChangeTextStyle_result.pptx", FileFormat.Pptx2007)
			Process.Start("ChangeTextStyle_result.pptx")
		End Sub
	End Class
End Namespace