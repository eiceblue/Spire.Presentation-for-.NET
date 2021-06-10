Imports System.Collections
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace GetSectionIndex
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

			Dim section As Section = ppt.SectionList(0)

			'Get the index of the section
			Dim index As Integer = ppt.SectionList.IndexOf(section)
			MessageBox.Show("The section index is: " & index)
		End Sub

	End Class
End Namespace