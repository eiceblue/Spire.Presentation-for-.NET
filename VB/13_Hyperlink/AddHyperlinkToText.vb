Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace AddHyperlinkToText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AddHyperlinkToText.pptx")

			'Find the text we want to add link to it.
			Dim shape As IAutoShape = TryCast(presentation.Slides(0).Shapes(0), IAutoShape)
			Dim tp As TextParagraph = shape.TextFrame.TextRange.Paragraph
			Dim temp As String = tp.Text

			'Split the original text.
			Dim textToLink As String = "Spire.Presentation"
			Dim strSplit() As String = temp.Split(New String() { "Spire.Presentation" }, StringSplitOptions.None)

			'Clear all text.
			tp.TextRanges.Clear()

			'Add new text.
			Dim tr As New TextRange(strSplit(0))
			tp.TextRanges.Append(tr)

			'Add the hyperlink.
			tr = New TextRange(textToLink)
			tr.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html"
			tp.TextRanges.Append(tr)

			Dim result As String = "Result-AddHyperlinkToText.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

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