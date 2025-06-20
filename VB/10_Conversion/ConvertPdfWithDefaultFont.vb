Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace ConvertPdfWithDefaultFont
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load PPT from disk
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ConvertPdfWithDefaultFont.pptx")

			'The font is preferred to convert to pdf or pictures, when the font used in the document is not installed in the system
			Presentation.SetDefaultFontName("Arial")

			'Save to file
			ppt.SaveToFile("ConvertPdfWithDefaultFont_out.pdf", FileFormat.PDF)

			'Launch and view the resulted file
			PresentationDocViewer("ConvertPdfWithDefaultFont_out.pdf")
		End Sub
		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
