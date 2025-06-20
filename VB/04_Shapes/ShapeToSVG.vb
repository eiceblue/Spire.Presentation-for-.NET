Imports Spire.Presentation
Imports System.IO

Namespace ShapeToSVG
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new Presentation object
			Dim ppt As New Presentation()

			' Load a PowerPoint file ("ShapeToSVG.pptx")
			ppt.LoadFromFile("..\..\..\..\..\..\..\Data\toSVG.pptx")

			' Access the first slide in the presentation
			Dim slide As ISlide = ppt.Slides(0)

			' Initialize a counter for file naming
			Dim num As Integer = 0

			' Iterate through each shape in the slide
			For Each shape As IShape In slide.Shapes
				' Save the shape as SVG format
				Dim svgByte() As Byte = shape.SaveAsSvg()

				' Create a new FileStream for writing the SVG content to a file
				Dim fs As New FileStream("shape_" & num & ".svg", FileMode.Create)

				' Write the SVG content to the file
				fs.Write(svgByte, 0, svgByte.Length)

				' Close the FileStream
				fs.Close()

				' Increment the counter for the next file naming
				num += 1
			Next shape

			' Dispose of the Presentation object to release resources
			ppt.Dispose()
		End Sub
	End Class
End Namespace
