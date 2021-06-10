Imports Spire.Presentation

Namespace MixFontStyles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\FontStyle.pptx")

			'Get the second shape of the first slide
			Dim shape As IAutoShape = TryCast(ppt.Slides(0).Shapes(1), IAutoShape)
			'Get the text from the shape 
			Dim originalText As String = shape.TextFrame.Text

			'Split the string by specified words and return substrings to a string array
			Dim splitArray() As String = originalText.Split(New String() { "bold", "red", "underlined", "bigger font size" }, StringSplitOptions.None)

			'Remove the paragraph from TextRange
			Dim tp As TextParagraph = shape.TextFrame.TextRange.Paragraph
			tp.TextRanges.Clear()

			'Append normal text that is in front of 'bold' to the paragraph
			Dim tr As New TextRange(splitArray(0))
			tp.TextRanges.Append(tr)
			'Set font style of the text 'bold' as bold
			tr = New TextRange("bold")
			tr.IsBold = TriState.True
			tp.TextRanges.Append(tr)

			'Append normal text that is in front of 'red' to the paragraph
			tr = New TextRange(splitArray(1))
			tp.TextRanges.Append(tr)
			'Set the color of the text 'red' as red
			tr = New TextRange("red")
			tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			tr.Format.Fill.SolidColor.Color = Color.Red
			tp.TextRanges.Append(tr)

			'Append normal text that is in front of 'underlined' to the paragraph
			tr = New TextRange(splitArray(2))
			tp.TextRanges.Append(tr)
			'Underline the text 'undelined'
			tr = New TextRange("underlined")
			tr.TextUnderlineType = TextUnderlineType.Single
			tp.TextRanges.Append(tr)

			'Append normal text that is in front of 'bigger font size' to the paragraph
			tr = New TextRange(splitArray(3))
			tp.TextRanges.Append(tr)
			'Set a large font for the text 'bigger font size'
			tr = New TextRange("bigger font size")
			tr.FontHeight = 35
			tp.TextRanges.Append(tr)

			'Append other normal text
			tr = New TextRange(splitArray(4))
			tp.TextRanges.Append(tr)

			'Save the document
			Dim result As String = "MixFontStyles.pptx"
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