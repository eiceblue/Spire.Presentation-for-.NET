Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace SetParagraphFont
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Az2.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'Access the first and second placeholder in the slide and typecasting it as AutoShape
			Dim tf1 As ITextFrameProperties = (CType(slide.Shapes(0), IAutoShape)).TextFrame
			Dim tf2 As ITextFrameProperties = (CType(slide.Shapes(1), IAutoShape)).TextFrame

			' Access the first Paragraph
			Dim para1 As TextParagraph = tf1.Paragraphs(0)
			Dim para2 As TextParagraph = tf2.Paragraphs(0)

			'Justify the paragraph
			para2.Alignment = TextAlignmentType.Justify

			'Access the first text range
			Dim textRange1 As TextRange = para1.FirstTextRange
			Dim textRange2 As TextRange = para2.FirstTextRange

			'Define new fonts
			Dim fd1 As New TextFont("Elephant")
			Dim fd2 As New TextFont("Castellar")

			' Assign new fonts to text range
			textRange1.LatinFont = fd1
			textRange2.LatinFont = fd2

			' Set font to Bold
			textRange1.Format.IsBold = TriState.True
			textRange2.Format.IsBold = TriState.False

			' Set font to Italic
			textRange1.Format.IsItalic = TriState.False
			textRange2.Format.IsItalic = TriState.True

			' Set font color
			textRange1.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange1.Fill.SolidColor.Color = Color.Purple
			textRange2.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange2.Fill.SolidColor.Color = Color.Peru

			Dim result As String = "SetParagraphFont_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)
			Viewer(result)
		End Sub

		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace