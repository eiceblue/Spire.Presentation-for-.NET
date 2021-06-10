Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation

Namespace InsertVideo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\InsertVideo.pptx")

			'Add title
			Dim rec_title As New RectangleF(50, 280, 160, 50)
			Dim shape_title As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec_title)
			shape_title.ShapeStyle.LineColor.Color = Color.Transparent

			shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			Dim para_title As New TextParagraph()
			para_title.Text = "Video:"
			para_title.Alignment = TextAlignmentType.Center
			para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
			para_title.TextRanges(0).FontHeight = 32
			para_title.TextRanges(0).IsBold = TriState.True
			para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			para_title.TextRanges(0).Fill.SolidColor.Color = Color.FromArgb(68, 68, 68)
			shape_title.TextFrame.Paragraphs.Append(para_title)

			'Insert video into the document
			Dim videoRect As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 125, 240, 150, 150)
			Dim video As IVideo = presentation.Slides(0).Shapes.AppendVideoMedia(Path.GetFullPath("..\..\..\..\..\..\Data\Video.mp4"), videoRect)
			video.PictureFill.Picture.Url = "..\..\..\..\..\..\Data\Video.png"

			'Save the document
			presentation.SaveToFile("video.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("video.pptx")
		End Sub
	End Class
End Namespace