Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetGlobalCustomFonts
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Set global custom fonts 
			Presentation.SetCustomFontsDirctory("..\..\..\..\..\..\Data\fonts")

			'Create a PPT document
			Dim ppt As New Presentation()

			'Load PPT file 
			ppt.LoadFromFile("..\..\..\..\..\..\Data\toPDF.pptx")

			'Save the PPT to PDF file format
			Dim result As String = "output.pdf"
			ppt.SaveToFile(result, FileFormat.PDF)

			Process.Start(result)
		End Sub
	End Class
End Namespace