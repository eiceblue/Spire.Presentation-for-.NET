Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace AddLineWithArrow
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			ppt.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			ppt.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Add a line to the slides and set its color to red
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Line, New RectangleF(150, 100, 100, 100))
			shape.ShapeStyle.LineColor.Color = Color.Red
			'Set the line end type as StealthArrow
			shape.Line.LineEndType = LineEndType.StealthArrow

			'Add a line to the slides and use default color
			shape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Line, New RectangleF(300, 150, 100, 100))
			shape.Rotation = -45
			'Set the line end type as TriangleArrowHead
			shape.Line.LineEndType = LineEndType.TriangleArrowHead

			'Add a line to the slides and set its color to Green
			shape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Line, New RectangleF(450, 100, 100, 100))
			shape.ShapeStyle.LineColor.Color = Color.Green
			shape.Rotation = 90
			'Set the line begin type as TriangleArrowHead
			shape.Line.LineBeginType = LineEndType.StealthArrow

			'Save the document
			Dim result As String = "AddLineWithArrow.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace