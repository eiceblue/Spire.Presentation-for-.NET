Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace FindShapeByAltText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\FindShapeByAltText.pptx")

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'Find shape in the slide
			Dim shape As IShape = FindShape(slide, "Shape1")

			If shape IsNot Nothing Then
				MessageBox.Show(shape.Name)
			End If

		End Sub
		Private Function FindShape(ByVal slide As ISlide, ByVal altText As String) As IShape
			'Loop through shapes in the slide
			For Each shape As IShape In slide.Shapes
				'Find the shape whose alternative text is altText
				If shape.AlternativeText.CompareTo(altText) = 0 Then
					Return shape
				End If
			Next shape
			Return Nothing
		End Function
	End Class
End Namespace