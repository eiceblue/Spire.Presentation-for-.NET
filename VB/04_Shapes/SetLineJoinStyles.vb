Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetLineJoinStyles
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

			'Add three shapes
			Dim shape1 As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 150, 150, 50))
			Dim shape2 As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(250, 150, 150, 50))
			Dim shape3 As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(450, 150, 150, 50))

			'Fill shapes
			shape1.Fill.FillType = FillFormatType.Solid
			shape1.Fill.SolidColor.Color = Color.CadetBlue
			shape2.Fill.FillType = FillFormatType.Solid
			shape2.Fill.SolidColor.Color = Color.CadetBlue
			shape3.Fill.FillType = FillFormatType.Solid
			shape3.Fill.SolidColor.Color = Color.CadetBlue

			'Fill lines of shapes
			shape1.Line.FillType = FillFormatType.Solid
			shape1.Line.SolidFillColor.Color = Color.DarkGray
			shape2.Line.FillType = FillFormatType.Solid
			shape2.Line.SolidFillColor.Color = Color.DarkGray
			shape3.Line.FillType = FillFormatType.Solid
			shape3.Line.SolidFillColor.Color = Color.DarkGray

			'Set the line width
			shape1.Line.Width = 10
			shape2.Line.Width = 10
			shape3.Line.Width = 10

			'Set the join styles of lines
			shape1.Line.JoinStyle = LineJoinType.Bevel
			shape2.Line.JoinStyle = LineJoinType.Miter
			shape3.Line.JoinStyle = LineJoinType.Round

			'Add text in shapes
			shape1.TextFrame.Text = "Bevel Join Style"
			shape2.TextFrame.Text = "Miter Join Style"
			shape3.TextFrame.Text = "Round Join Style"

			'Save the document
			Dim result As String = "SetLineJoinStyles_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace