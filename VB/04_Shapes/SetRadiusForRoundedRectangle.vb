Imports Spire.Presentation

Namespace SetRadiusForRoundedRectangle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			'Create a PPT document
			Dim presentation As New Presentation()

			'Get the first slide 
			Dim slide As ISlide = presentation.Slides(0)

			'Insert a rectangle with four round corners and set its radius
			Dim shape1 As IAutoShape = slide.Shapes.AppendShape(ShapeType.RoundCornerRectangle, New RectangleF(50, 50, 150, 150))
			shape1.SetRoundRadius(shape1.Width \ 3)

			'Insert a rectangle with one round corner and set its radius
			Dim shape2 As IAutoShape = slide.Shapes.AppendShape(ShapeType.OneRoundCornerRectangle, New RectangleF(250, 50, 150, 150))
			shape2.SetRoundRadius(shape2.Width \ 3)

			'Insert a rectangle with one round corner and which one round cornet is snipped and set its radius
			Dim shape3 As IAutoShape = slide.Shapes.AppendShape(ShapeType.OneSnipOneRoundCornerRectangle, New RectangleF(450, 50, 150, 150))
			shape3.SetRoundRadius(shape3.Width \ 3)

			'Insert a rectangle with two diagonal round corners and set its radius
			Dim shape4 As IAutoShape = slide.Shapes.AppendShape(ShapeType.TwoDiagonalRoundCornerRectangle, New RectangleF(50, 250, 150, 150))
			shape4.SetRoundRadius(shape4.Width \ 3)

			'Insert a rectangle with two same side round corners and set its radius
			Dim shape5 As IAutoShape = slide.Shapes.AppendShape(ShapeType.TwoSamesideRoundCornerRectangle, New RectangleF(250, 250, 150, 150))
			shape5.SetRoundRadius(shape5.Width \ 3)


			'Save to file.
			Dim result As String = "output.pptx"
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