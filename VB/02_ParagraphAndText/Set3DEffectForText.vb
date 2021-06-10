Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace Set3DEffectForText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a new presentation object
			Dim ppt As New Presentation()

			'Get the first slide
			Dim slide As ISlide = ppt.Slides(0)

			'Append a new shape to slide and set the line color and fill type
			Dim shape As IAutoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(30, 40, 650, 200))
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Fill.FillType = FillFormatType.None

			'Add text to the shape
			shape.AppendTextFrame("This demo shows how to add 3D effect text to Presentation slide")

			'Set the color of text in shape
			Dim textRange As TextRange = shape.TextFrame.TextRange
			textRange.Fill.FillType = FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.LightBlue

			'Set the Font of text in shape
			textRange.FontHeight = 40
			textRange.LatinFont = New TextFont("Gulim")

			'Set 3D effect for text
			shape.TextFrame.TextThreeD.ShapeThreeD.PresetMaterial = PresetMaterialType.Matte
			shape.TextFrame.TextThreeD.LightRig.PresetType = PresetLightRigType.Sunrise
			shape.TextFrame.TextThreeD.ShapeThreeD.TopBevel.PresetType = BevelPresetType.Circle
			shape.TextFrame.TextThreeD.ShapeThreeD.ContourColor.Color = Color.Green
			shape.TextFrame.TextThreeD.ShapeThreeD.ContourWidth = 3

			'Save the document
			Dim result As String = "Set3DEffectForText.pptx"
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