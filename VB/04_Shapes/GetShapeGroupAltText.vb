Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.IO

Namespace GetShapeGroupAltText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\GetShapeGroupAltText.pptx")

			Dim builder As New StringBuilder()

			'Loop through slides and shapes
			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					If TypeOf shape Is GroupShape Then
						'Find the shape group
						Dim groupShape As GroupShape = TryCast(shape, GroupShape)
						For Each gShape As IShape In groupShape.Shapes
							'Append the alternative text in builder
							builder.AppendLine(gShape.AlternativeText)
						Next gShape
					End If
				Next shape
			Next slide

			'Write the content in txt file
			Dim output As String="GetShapeAltText_result.txt"
			File.WriteAllText(output, builder.ToString())

			'Launch the txt file
			OutputViewer(output)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace