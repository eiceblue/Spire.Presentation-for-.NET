Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace SuperscriptAndSubscript
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Az.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			'Add a shape 
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(150, 100, 200, 50))
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			shape.TextFrame.Paragraphs.Clear()

			shape.AppendTextFrame("Test")
			Dim tr As New TextRange("superscript")
			shape.TextFrame.Paragraphs(0).TextRanges.Append(tr)

			'Set superscript text
			shape.TextFrame.Paragraphs(0).TextRanges(1).Format.ScriptDistance = 30

			Dim textRange As TextRange = shape.TextFrame.Paragraphs(0).TextRanges(0)
			textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.Black
			textRange.FontHeight = 20
			textRange.LatinFont = New TextFont("Lucida Sans Unicode")

			textRange = shape.TextFrame.Paragraphs(0).TextRanges(1)
			textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.CadetBlue
			textRange.LatinFont = New TextFont("Lucida Sans Unicode")


			shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(150, 150, 200, 50))
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			shape.TextFrame.Paragraphs.Clear()

			shape.AppendTextFrame("Test")
			tr = New TextRange("subscript")
			shape.TextFrame.Paragraphs(0).TextRanges.Append(tr)

			'Set subscript text
			shape.TextFrame.Paragraphs(0).TextRanges(1).Format.ScriptDistance = -25

			textRange = shape.TextFrame.Paragraphs(0).TextRanges(0)
			textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.Black
			textRange.FontHeight = 20
			textRange.LatinFont = New TextFont("Lucida Sans Unicode")

			textRange = shape.TextFrame.Paragraphs(0).TextRanges(1)
			textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.CadetBlue
			textRange.LatinFont = New TextFont("Lucida Sans Unicode")


			Dim result As String = "SuperscriptAndSubscript_result.pptx"
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