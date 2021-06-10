Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace Background
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Set background Image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\backgroundImg.png"
			Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
			presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)

			'Add title
			Dim rec_title As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 200, 70, 380, 50)
			Dim shape_title As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec_title)
			shape_title.Line.FillType = FillFormatType.None
			shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			Dim para_title As New TextParagraph()
			para_title.Text = "Background Sample"
			para_title.Alignment = TextAlignmentType.Center
			para_title.TextRanges(0).LatinFont = New TextFont("Lucida Sans Unicode")
			para_title.TextRanges(0).FontHeight = 36
			para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			para_title.TextRanges(0).Fill.SolidColor.Color = Color.DarkSlateBlue
			shape_title.TextFrame.Paragraphs.Append(para_title)

			'Add new shape to PPT document
			Dim rec As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 300, 155, 600, 200)
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)
			shape.Line.FillType = FillFormatType.None
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

			Dim para As New TextParagraph()
			para.Text = "Spire.Presentation for .NET support PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc."

			para.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			para.TextRanges(0).Fill.SolidColor.Color = Color.CadetBlue
			para.TextRanges(0).FontHeight = 26
			shape.TextFrame.Paragraphs.Append(para)

			'Save the document
			presentation.SaveToFile("Background_result.pptx", FileFormat.Pptx2010)
			Process.Start("Background_result.pptx")
		End Sub
	End Class
End Namespace