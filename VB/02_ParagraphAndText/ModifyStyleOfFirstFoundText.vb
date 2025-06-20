Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace ModifyStyleOfFirstFoundText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\TextTemplate.pptx")

			'Find first "Spire"
			Dim text As String = "Spire"
			Dim textRange As TextRange = ppt.Slides(0).FindFirstTextAsRange(text)

			'Modify the style
			textRange.Fill.FillType = FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.Red
			textRange.FontHeight = 28
			textRange.LatinFont = New TextFont("Calibri")
			textRange.IsBold = TriState.True
			textRange.IsItalic = TriState.True
			textRange.TextUnderlineType = TextUnderlineType.Double
			textRange.TextStrikethroughType = TextStrikethroughType.Single

			'Save the document
			Dim result As String = "Result.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			PresentationDocViewer(result)
		End Sub

	Private Shared Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
	End Sub

	End Class
End Namespace