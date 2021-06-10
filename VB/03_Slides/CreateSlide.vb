Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace CreateSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Add new slide
			presentation.Slides.Append()

			'Set the background image
			For i As Integer = 0 To 1
				Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
				Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
				presentation.Slides(i).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
				presentation.Slides(i).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite
			Next i

			'Add title
			Dim rec_title As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 200, 70, 400, 50)
			Dim shape_title As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec_title)
			shape_title.ShapeStyle.LineColor.Color = Color.White
			shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			Dim para_title As New TextParagraph()
			para_title.Text = "E-iceblue"
			para_title.Alignment = TextAlignmentType.Center
			para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
			para_title.TextRanges(0).FontHeight = 36
			para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			para_title.TextRanges(0).Fill.SolidColor.Color = Color.Black
			shape_title.TextFrame.Paragraphs.Append(para_title)

			'Append new shape
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 150, 600, 280))
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			shape.Line.FillType = FillFormatType.None
			'Add text to shape
			shape.AppendTextFrame("Welcome to use Spire.Presentation for .NET.")

			'Add new paragraph
			Dim pare As New TextParagraph()
			pare.Text = ""
			shape.TextFrame.Paragraphs.Append(pare)

			'Add new paragraph
			pare = New TextParagraph()
			pare.Text = "Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine."
			shape.TextFrame.Paragraphs.Append(pare)

			'Set the Font
			For Each para As TextParagraph In shape.TextFrame.Paragraphs
				para.TextRanges(0).LatinFont = New TextFont("Myriad Pro")
				para.TextRanges(0).FontHeight = 24
				para.TextRanges(0).Fill.FillType = FillFormatType.Solid
				para.TextRanges(0).Fill.SolidColor.Color = Color.Black
				para.Alignment = TextAlignmentType.Left
			Next para

			'Append new shape - SixPointedStar
			shape = presentation.Slides(1).Shapes.AppendShape(ShapeType.SixPointedStar, New RectangleF(100, 100, 100, 100))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.Orange
			shape.ShapeStyle.LineColor.Color = Color.White

			'Append new shape
			shape = presentation.Slides(1).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 250, 600, 50))
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

			'Add text to shape
			shape.AppendTextFrame("This is newly added Slide.")

			'Set the Font
			shape.TextFrame.Paragraphs(0).TextRanges(0).LatinFont = New TextFont("Myriad Pro")
			shape.TextFrame.Paragraphs(0).TextRanges(0).FontHeight = 24
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.FillType = FillFormatType.Solid
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.SolidColor.Color = Color.Black
			shape.TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Left
			shape.TextFrame.Paragraphs(0).Indent = 35

			'Save the document
			presentation.SaveToFile("CreateSlide.pptx", FileFormat.Pptx2010)
			Process.Start("CreateSlide.pptx")
		End Sub
	End Class
End Namespace