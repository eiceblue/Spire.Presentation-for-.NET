Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace GetTextLinesInShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\GetLinesInShape.pptx")

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

		   Dim sb As New StringBuilder()

			' Iterate the shapes in the slide
			For i As Integer = 0 To slide.Shapes.Count - 1
				' Get shape 
				Dim shape As IAutoShape = CType(slide.Shapes(i), IAutoShape)
				sb.Append("shape" & i & ":" & vbCrLf)

				' Get text lines in the shape and get the text
				Dim lines As IList(Of LineText) = shape.TextFrame.GetLayoutLines()
				For j As Integer = 0 To lines.Count - 1
					sb.Append("line[" & j & "]:" & lines(j).Text & vbCrLf)
				Next j
			Next i

			File.WriteAllText("GetLinesInShape.txt", sb.ToString())

			Process.Start("GetLinesInShape.txt")

			presentation.Dispose()
		End Sub
	End Class
End Namespace