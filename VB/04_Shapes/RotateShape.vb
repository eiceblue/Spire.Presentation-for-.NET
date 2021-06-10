Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace RotateShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\RotateShape.pptx")

			'Get the shapes 
			Dim shape As IAutoShape = TryCast(ppt.Slides(0).Shapes(0), IAutoShape)

			'Set the rotation
			shape.Rotation =60

			TryCast(ppt.Slides(0).Shapes(1), IAutoShape).Rotation = 120
			TryCast(ppt.Slides(0).Shapes(2), IAutoShape).Rotation = 180
			TryCast(ppt.Slides(0).Shapes(3), IAutoShape).Rotation = 240

			'Save the document
			ppt.SaveToFile("RotateShape_result.pptx", FileFormat.Pptx2010)
			Process.Start("RotateShape_result.pptx")
		End Sub

	End Class
End Namespace
