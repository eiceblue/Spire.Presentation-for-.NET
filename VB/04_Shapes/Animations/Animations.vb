Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing.Animation
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace Animations
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\Animations.pptx")

			'Add title
			Dim rec_title As New RectangleF(50, 200, 200, 50)
			Dim shape_title As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec_title)
			shape_title.ShapeStyle.LineColor.Color = Color.Transparent

			shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			Dim para_title As New TextParagraph()
			para_title.Text = "Animations:"
			para_title.Alignment = TextAlignmentType.Center
			para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
			para_title.TextRanges(0).FontHeight = 32
			para_title.TextRanges(0).IsBold = TriState.True
			para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			para_title.TextRanges(0).Fill.SolidColor.Color = Color.FromArgb(68, 68, 68)
			shape_title.TextFrame.Paragraphs.Append(para_title)

			'Set the animation of slide to Circle
			presentation.Slides(0).SlideShowTransition.Type = Spire.Presentation.Drawing.Transition.TransitionType.Circle

			'Append new shape - Triangle
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Triangle, New RectangleF(100, 280, 80, 80))

			'Set the color of shape
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.CadetBlue
			shape.ShapeStyle.LineColor.Color = Color.White

			'Set the animation of shape
			shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Path4PointStar)

			'Append new shape - Rectangle and set animation
			shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(210, 280, 150, 80))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.CadetBlue
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.AppendTextFrame("Animated Shape")
			shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel)

			'Append new shape - Cloud and set the animation
			shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Cloud, New RectangleF(390, 280, 80, 80))
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.White
			shape.ShapeStyle.LineColor.Color = Color.CadetBlue
			shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedZoom)

			'Save the document
			presentation.SaveToFile("animations.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("animations.pptx")

		End Sub
	End Class
End Namespace