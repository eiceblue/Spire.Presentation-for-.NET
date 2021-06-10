Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace Set3DEffectForShape
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

			'Add shape1 and fill it with color
			Dim shape1 As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.RoundCornerRectangle, New RectangleF(150, 150, 150, 150))
			shape1.Fill.FillType = FillFormatType.Solid
			shape1.Fill.SolidColor.KnownColor = KnownColors.SkyBlue
			'Initialize a new instance of the 3-D class for shape1 and set its properties
			Dim effect1 As ShapeThreeD = shape1.ThreeD.ShapeThreeD
			effect1.PresetMaterial = PresetMaterialType.Powder
			effect1.TopBevel.PresetType = BevelPresetType.ArtDeco
			effect1.TopBevel.Height = 4
			effect1.TopBevel.Width = 12
			effect1.BevelColorMode = BevelColorType.Contour
			effect1.ContourColor.KnownColor = KnownColors.LightBlue
			effect1.ContourWidth = 3.5

			'Add shape2 and fill it with color
			Dim shape2 As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Pentagon, New RectangleF(400, 150, 150, 150))
			shape2.Fill.FillType = FillFormatType.Solid
			shape2.Fill.SolidColor.KnownColor = KnownColors.LightGreen
			'Initialize a new instance of the 3-D class for shape2 and set its properties
			Dim effect2 As ShapeThreeD = shape2.ThreeD.ShapeThreeD
			effect2.PresetMaterial = PresetMaterialType.SoftEdge
			effect2.TopBevel.PresetType = BevelPresetType.SoftRound
			effect2.TopBevel.Height = 12
			effect2.TopBevel.Width = 12
			effect2.BevelColorMode = BevelColorType.Contour
			effect2.ContourColor.KnownColor = KnownColors.LawnGreen
			effect2.ContourWidth = 5

			'Save the document
			Dim result As String = "Set3DEffectForShape.pptx"
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