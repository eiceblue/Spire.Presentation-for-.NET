Imports System.Collections
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace UngroupShapes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()
			'Load the document from disk
			ppt.LoadFromFile("..\..\..\..\..\..\Data\GroupShapes.pptx")
			'Get the GroupShape
			Dim groupShape As GroupShape = TryCast(ppt.Slides(0).Shapes(0), GroupShape)
			'Ungroup the shapes
			ppt.Slides(0).Ungroup(groupShape)
			'Save the document
			Dim result As String = "UngroupShapes.pptx"
			ppt.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013)
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