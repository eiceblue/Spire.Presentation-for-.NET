Imports Spire.Presentation

Namespace SetTextDirection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Append a shape with text to the first slide
			Dim textboxShape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(250, 70, 100, 400))
			textboxShape.ShapeStyle.LineColor.Color = Color.Transparent
			textboxShape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textboxShape.Fill.SolidColor.Color = Color.LightBlue
			textboxShape.TextFrame.Text = "You Are Welcome Here"
			'Set the text direction to vertical
			textboxShape.TextFrame.VerticalTextType = VerticalTextType.Vertical

			'Append another shape with text to the slide
			textboxShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(350, 70, 100, 400))
			textboxShape.ShapeStyle.LineColor.Color = Color.Transparent
			textboxShape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textboxShape.Fill.SolidColor.Color = Color.LightGray
			'Append some asian characters
			textboxShape.TextFrame.Text = "ª∂”≠π‚¡Ÿ"
			'Set the VerticalTextType as EastAsianVertical to aviod rotating text 90 degrees
			textboxShape.TextFrame.VerticalTextType = VerticalTextType.EastAsianVertical

			'Save the document
			Dim result As String = "SetTextDirection.pptx"
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