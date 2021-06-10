Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace DetectUsedThemes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\Themes.pptx")

			Dim sb As New StringBuilder()
			Dim themeName As String = Nothing
			sb.AppendLine("This is the name list of the used theme below.")
			'Get the theme name of each slide in the document
			For Each slide As ISlide In ppt.Slides
				themeName = slide.Theme.Name
				sb.AppendLine(themeName)
			Next slide

			'Save to the text document
			Dim result As String = "DetectUsedThemes.txt"
			File.WriteAllText(result, sb.ToString())
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