Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetFormatForLines
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

			'Add a rectangle shape to the slide
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(100, 150, 200, 100))
			'Set the fill color of the rectangle shape
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.White
			'Apply some formatting on the line of the rectangle
			shape.Line.Style = TextLineStyle.ThickThin
			shape.Line.Width = 5
			shape.Line.DashStyle = LineDashStyleType.Dash
			'Set the color of the line of the rectangle
			shape.ShapeStyle.LineColor.Color = Color.SkyBlue

			'Add a ellipse shape to the slide
			shape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Ellipse, New RectangleF(400, 150, 200, 100))
			'Set the fill color of the ellipse shape
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.White
			'Apply some formatting on the line of the ellipse
			shape.Line.Style = TextLineStyle.ThickBetweenThin
			shape.Line.Width = 5
			shape.Line.DashStyle = LineDashStyleType.DashDot
			'Set the color of the line of the ellipse
			shape.ShapeStyle.LineColor.Color = Color.OrangeRed

			'Save the document
			Dim result As String = "SetFormatForLines.pptx"
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