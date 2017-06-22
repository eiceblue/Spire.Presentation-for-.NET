Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text
'Imports System.Threading.Tasks

Namespace FillShapeWithGradient
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation
			Dim ppt As New Presentation()
			'Add a rectangle to the slide
			Dim GradientShape As IAutoShape = CType(ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(200, 100, 287, 100)), IAutoShape)

			'Set the Fill Type of the Shape and color to Gradient
			GradientShape.Fill.FillType = FillFormatType.Gradient
			GradientShape.Fill.Gradient.GradientStops.Append(0, Color.Purple)
			GradientShape.Fill.Gradient.GradientStops.Append(1, Color.Red)

			ppt.SaveToFile("CreateGradientShape.pptx", FileFormat.Pptx2010)
			Process.Start("CreateGradientShape.pptx")
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
