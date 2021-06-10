Imports System.Collections
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace GroupShapes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Get the first slide
			Dim slide As ISlide = ppt.Slides(0)

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			slide.Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Create two shapes in the slide
			Dim rectangle As IShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(250, 180, 200, 40))
			rectangle.Fill.FillType = FillFormatType.Solid
			rectangle.Fill.SolidColor.KnownColor = KnownColors.SkyBlue
			rectangle.Line.Width = 0.1f
			Dim ribbon As IShape = slide.Shapes.AppendShape(ShapeType.Ribbon2, New RectangleF(290, 155, 120, 80))
			ribbon.Fill.FillType = FillFormatType.Solid
			ribbon.Fill.SolidColor.KnownColor = KnownColors.LightPink
			ribbon.Line.Width = 0.1f

			'Add the two shape objects to an array list
			Dim list As New ArrayList()
			list.Add(rectangle)
			list.Add(ribbon)

			'Group the shapes in the list
			ppt.Slides(0).GroupShapes(list)

			'Save the document
			Dim result As String = "GroupShapes.pptx"
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