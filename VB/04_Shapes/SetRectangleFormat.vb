Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetRectangleFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Add a shape
			Dim rect As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 100, 100, 200, 100)
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rect)

			'Set the fill format of shape
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.CadetBlue

			'Set the fill format of line
			shape.Line.FillType = FillFormatType.Solid
			shape.Line.SolidFillColor.Color = Color.DimGray

			'Save the document
			Dim result As String = "SetRectangleFormat_result.pptx"
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