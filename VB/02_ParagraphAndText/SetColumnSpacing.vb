Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetColumnSpacing
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new PPT
			Dim presentation As New Presentation()

			' Append a shape in the first slide
			Dim slide As ISlide = presentation.Slides(0)
			Dim shape As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 70, 600, 400))
			shape.TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Left
			shape.Fill.FillType = FillFormatType.None

			' Set column and column spacing
			shape.TextFrame.ColumnCount = 2
			shape.TextFrame.ColumnSpacing = 20.50f
			' Append text
			shape.TextFrame.Text = vbCrLf & "Spire.Presentation for .NET is a professional PowerPoint® compatible API that enables developers to create, read, write, modify, convert and Print PowerPoint documents on any .NET platform (Target .NET Framework, .NET Core, .NET Standard, .NET 5.0, .NET 6.0, Xamarin & Mono Android). As an independent PowerPoint .NET API, Spire.Presentation for .NET doesn't need Microsoft PowerPoint to be installed on machines." & vbCrLf & vbCrLf & vbCrLf & "Spire.Presentation for .NET supports PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also supports exporting presentation slides to images (PNG, JPG, TIFF, EMF, SVG), PDF, XPS, HTML format etc."
			For Each paragraph As TextParagraph In shape.TextFrame.Paragraphs
				For Each textRange As TextRange In paragraph.TextRanges
					' Set font for text
					textRange.Fill.FillType = FillFormatType.Solid
					textRange.Fill.SolidColor.Color = Color.Black
					textRange.FontHeight = 16
					textRange.LatinFont = New TextFont("Open Sans")
				Next textRange
			Next paragraph

			Dim result As String = "Result-SetColumnSpacing.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)
			' Dispose the document
			presentation.Dispose()

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