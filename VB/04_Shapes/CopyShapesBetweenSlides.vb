Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace CopyShapesBetweenSlides
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load the sample document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\CopyShapesBetweenSlides.pptx")

			'Define the source slide and target slide
			Dim sourceSlide As ISlide = ppt.Slides(0)
			Dim targetSlide As ISlide = ppt.Slides(1)

			'Copy the first shape from the source slide to the target slide
			targetSlide.Shapes.AddShape(CType(sourceSlide.Shapes(0), Shape))

			Dim result As String = "CopyShapesBetweenSlides-result.pptx"
			'Save the document to file 
			ppt.SaveToFile(result, FileFormat.Pptx2013)

			Viewer(result)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace