Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace GetTextStyleEffectiveData
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Az1.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			'Get a shape 
			Dim shape As IAutoShape = TryCast(presentation.Slides(0).Shapes(0), IAutoShape)

			Dim str As New StringBuilder()
			For p As Integer = 0 To shape.TextFrame.Paragraphs.Count - 1
				Dim paragraph = shape.TextFrame.Paragraphs(p)
				str.AppendLine("Text style for Paragraph " & p & " :")
				'Get the paragraph style
				str.AppendLine(" Indent: " & paragraph.Indent)
				str.AppendLine(" Alignment: " & paragraph.Alignment)
				str.AppendLine(" Font alignment: " & paragraph.FontAlignment)
				str.AppendLine(" Hanging punctuation: " & paragraph.HangingPunctuation)
				str.AppendLine(" Line spacing: " & paragraph.LineSpacing)
				str.AppendLine(" Space before: " & paragraph.SpaceBefore)
				str.AppendLine(" Space after: " & paragraph.SpaceAfter.ToString())
				str.AppendLine()
				For r As Integer = 0 To paragraph.TextRanges.Count - 1
					Dim textRange = paragraph.TextRanges(r)
					str.AppendLine("  Text style for Paragraph " & p & " TextRange " & r & " :")
					'Get the text range style
					str.AppendLine("    Font height: " & textRange.FontHeight)
					str.AppendLine("    Language: " & textRange.Language)
					str.AppendLine("    Font: " & textRange.LatinFont.FontName)
					str.AppendLine()
				Next r
			Next p

			Dim result As String = "GetTextStyleEffectiveData_result.txt"
			File.WriteAllText(result, str.ToString())
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