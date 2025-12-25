Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace AddSVGAsShapes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Add the SVG file as a shape onto the first slide of the presentation
			presentation.Slides(0).Shapes.AddFromSVGAsShapes("..\..\..\..\..\..\Data\AddSVGAsShapes.svg")

			'Save the document
			Dim result As String = "AddSVGAsShapes.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Dispose the presentation object
			presentation.Dispose()

			'Launch the file
			Process.Start(result)
		End Sub
	End Class
End Namespace