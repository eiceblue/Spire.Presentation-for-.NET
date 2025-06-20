Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace GetTextFrameSizeOfShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\GetTextFrameSizeOfShape.pptx")

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			Dim sb As New StringBuilder()

			' Iterate the shapes in the slide
			For i As Integer = 0 To slide.Shapes.Count - 1

				Dim autoShape As IAutoShape = TryCast(slide.Shapes(i), IAutoShape)
				Dim size As SizeF = autoShape.TextFrame.GetTextSize()
				sb.AppendLine("The size of text frame in shape" & i & ", width:" & size.Width & " height:" & size.Height)
			Next i

			File.WriteAllText("GetTextFrameSizeOfShape.txt", sb.ToString())

			Process.Start("GetTextFrameSizeOfShape.txt")

			presentation.Dispose()
		End Sub
	End Class
End Namespace