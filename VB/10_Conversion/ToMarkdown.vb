Imports Spire.Presentation

Namespace ToMarkdown

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create and load the file 
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ExtractText.pptx")
			' Convert to markdown format
			ppt.SaveToFile("ToMarkdown.md", FileFormat.Markdown)
			ppt.Dispose()

		End Sub
	End Class
End Namespace