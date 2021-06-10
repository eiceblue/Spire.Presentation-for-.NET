Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace AddSlideToSection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim savePath As String = "AddSlideToSection.pptx"
			Dim input As String = "..\..\..\..\..\..\Data\Section.pptx"

			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile(input)

			'Add a new shape to the PPT document
			presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(200, 50, 300, 100))

			'Create a new section and copy the first slide to it
			Dim NewSection As Section = presentation.SectionList.Append("New Section")
			NewSection.Insert(0, presentation.Slides(0))

			presentation.SaveToFile(savePath, FileFormat.Pptx2013)
			Process.Start(savePath)

		End Sub
	End Class
End Namespace