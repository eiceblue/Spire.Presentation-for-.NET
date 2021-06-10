Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetShadowEffectForText
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

			'Get reference of the slide
			Dim slide As ISlide = ppt.Slides(0)

			'Add a new rectangle shape to the first slide
			Dim shape As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(120, 100, 450, 200))
			shape.Fill.FillType = FillFormatType.None

			'Add the text to the shape and set the font for the text
			shape.AppendTextFrame("Text shading on slides")
			shape.TextFrame.Paragraphs(0).TextRanges(0).LatinFont = New TextFont("Arial Black")
			shape.TextFrame.Paragraphs(0).TextRanges(0).FontHeight = 21
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.FillType = FillFormatType.Solid
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.SolidColor.Color = Color.Black

			'//Add inner shadow and set all necessary parameters
			'InnerShadowEffect Shadow = InnerShadowEffect();

			'Add outer shadow and set all necessary parameters
			Dim Shadow As New OuterShadowEffect()

			Shadow.BlurRadius = 0
			Shadow.Direction = 50
			Shadow.Distance = 10
			Shadow.ColorFormat.Color = Color.LightBlue

			'shape.TextFrame.TextRange.EffectDag.InnerShadowEffect = Shadow;
			shape.TextFrame.TextRange.EffectDag.OuterShadowEffect = Shadow

			'Save the document
			Dim result As String = "SetShadowEffect.pptx"
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