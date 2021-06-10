Imports System.Collections
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace AddSection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\BlankSample.pptx")

			'Get the second slide
			Dim slide As ISlide = ppt.Slides(1)

			'Append section with section name at the end
			ppt.SectionList.Append("E-iceblue01")
			'Add section with slide
			ppt.SectionList.Add("section1", slide)

			Dim result As String = "AddSection.pptx"
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