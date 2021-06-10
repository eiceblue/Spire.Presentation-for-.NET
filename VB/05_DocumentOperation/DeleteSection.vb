Imports System.Collections
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace DeleteSection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()

			'Load the document from disk
			ppt.LoadFromFile("..\..\..\..\..\..\Data\AddSection.pptx")

			'//remove the specified section
			'ppt.SectionList.RemoveAt(3);
			'remove all sections
			ppt.SectionList.RemoveAll()

			'Save the document
			Dim result As String = "DeleteOneSelection.pptx"
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