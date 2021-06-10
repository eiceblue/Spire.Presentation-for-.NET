Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace PreventOrAllowChangingShape
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
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 100, 400, 150))

			'Set the shape format
			shape.Fill.FillType = FillFormatType.None
			shape.ShapeStyle.LineColor.Color = Color.LightBlue
			shape.TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Justify
			shape.TextFrame.Text = "Demo for locking shapes:" & vbLf & "    Green/Black stands for editable." & vbLf & "    Grey stands for non-editable."
			shape.TextFrame.Paragraphs(0).TextRanges(0).LatinFont = New TextFont("Arial Rounded MT Bold")
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.FillType = FillFormatType.Solid
			shape.TextFrame.Paragraphs(0).TextRanges(0).Fill.SolidColor.Color = Color.Black

			'The changes of selection and rotation are allowed
			shape.Locking.RotationProtection = False
			shape.Locking.SelectionProtection = False
			'The changes of size, position, shape type, aspect ratio, text editing and ajust handles are not allowed 
			shape.Locking.ResizeProtection = True
			shape.Locking.PositionProtection = True
			shape.Locking.ShapeTypeProtection = True
			shape.Locking.AspectRatioProtection = True
			shape.Locking.TextEditingProtection = True
			shape.Locking.AdjustHandlesProtection = True

			'Save the document
			Dim result As String = "PreventOrAllowChangingShape.pptx"
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