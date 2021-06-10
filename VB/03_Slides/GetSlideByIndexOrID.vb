Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace GetSlideByIndexOrID
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			 'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\BlankSample_N.pptx")

			'Get slide by index 0
			Dim slide1 As ISlide = presentation.Slides(0)
			'Append a shape in the slide
			Dim shape1 As IAutoShape=slide1.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(100, 100, 200, 100))
			'Add text in the shape
			shape1.TextFrame.Text = "Get slide by index"

			'Get slide by slide ID
			Dim slide2 As ISlide = presentation.FindSlide(CInt(Fix(presentation.Slides(1).SlideID)))
			'Append a shape in the slide
			Dim shape2 As IAutoShape = slide2.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(100, 100, 200, 100))
			'Add text in the shape
			shape2.TextFrame.Text = "Get slide by slide id"

			'Save the document
			Dim result As String = "GetSlideByIndexOrID_result.pptx"
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