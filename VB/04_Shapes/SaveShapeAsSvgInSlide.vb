Imports System.IO
Imports Spire.Presentation

Namespace SaveShapeAsSvgInSlide
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_7.pptx")

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			' Iterate the shapes in the slide
			For i As Integer = 0 To slide.Shapes.Count - 1
				' Save the shapes
				Dim svgByte() As Byte = slide.Shapes(i).SaveAsSvgInSlide()
				Dim fs As New FileStream("shapePath_" & i & ".svg", FileMode.Create)

				' Close the stream
				fs.Write(svgByte, 0, svgByte.Length)
				fs.Close()
			Next i

			' Dispose the Presentation object
			presentation.Dispose()
		End Sub
	End Class
End Namespace