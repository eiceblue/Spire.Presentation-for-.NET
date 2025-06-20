Imports Spire.Presentation

Namespace AddSlideUsingMasterLayout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AppendSlideWithMasterLayout.pptx")

			'Get Master layouts
			Dim iLayout As ILayout = presentation.Masters(0).Layouts(0)

			'Append new slide
			presentation.Slides.Append(iLayout)

			'Insert new slide
			presentation.Slides.Insert(1, iLayout)

			'Save to file.
			Dim result As String = "output.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file.
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace