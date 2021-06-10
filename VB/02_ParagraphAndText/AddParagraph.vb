Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace AddParagraph
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			ppt.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			ppt.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Append a new shape
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 70, 620, 150))
			shape.Fill.FillType = FillFormatType.None
			shape.ShapeStyle.LineColor.Color = Color.White

			'Set the alignment of paragraph
			shape.TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Left
			'Set the indent of paragraph
			shape.TextFrame.Paragraphs(0).Indent = 50
			'Set the linespacing of paragraph
			shape.TextFrame.Paragraphs(0).LineSpacing = 150
			'Set the text of paragraph
			shape.TextFrame.Text = "This powerful component suite contains the most up-to-date versions of all .NET WPF Silverlight components offered by E-iceblue."

			'Set the Font
			shape.TextFrame.Paragraphs(0).TextRanges(0).LatinFont = New TextFont("Arial Rounded MT Bold")
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.FillType = FillFormatType.Solid
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.SolidColor.Color = Color.Black

			'Save and launch the document
			Dim result As String = "AddParagraph.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			PresentationDocViewer(result)
		End Sub

	Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
	End Sub
	End Class
End Namespace