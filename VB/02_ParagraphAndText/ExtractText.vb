Imports Spire.Presentation
Imports System.ComponentModel
Imports System.IO
Imports System.Text

Namespace ExtractText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create a PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\edit.pptx")

			Dim sb As New StringBuilder()
			'foreach the slide and extract text
			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					If TypeOf shape Is IAutoShape Then
						For Each tp As TextParagraph In (TryCast(shape, IAutoShape)).TextFrame.Paragraphs
							sb.Append(tp.Text & Environment.NewLine)
						Next tp
					End If

				Next shape
			Next slide
			File.WriteAllText("Extract.txt", sb.ToString())
			Process.Start("Extract.txt")
		End Sub
	End Class
End Namespace
