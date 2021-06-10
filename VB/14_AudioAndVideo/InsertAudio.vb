Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation

Namespace InsertAudio
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\InsertAudio.pptx")

			'Add title
			Dim rec_title As New RectangleF(50, 240, 160,50)
			Dim shape_title As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec_title)
			shape_title.ShapeStyle.LineColor.Color = Color.Transparent

			shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			Dim para_title As New TextParagraph()
			para_title.Text = "Audio:"
			para_title.Alignment = TextAlignmentType.Center
			para_title.TextRanges(0).LatinFont = New TextFont("Myriad Pro Light")
			para_title.TextRanges(0).FontHeight = 32
			para_title.TextRanges(0).IsBold = TriState.True
			para_title.TextRanges(0).Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			para_title.TextRanges(0).Fill.SolidColor.Color = Color.FromArgb(68,68,68)
			shape_title.TextFrame.Paragraphs.Append(para_title)

			'Insert audio into the document
			Dim audioRect As New RectangleF(220, 240, 80, 80)
			presentation.Slides(0).Shapes.AppendAudioMedia(Path.GetFullPath("..\..\..\..\..\..\Data\Music.wav"), audioRect)

			'Save the document
			presentation.SaveToFile("Audio.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("Audio.pptx")
		End Sub
	End Class
End Namespace