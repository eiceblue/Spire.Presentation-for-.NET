Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ConvertUnhiddenSlidesToPdf
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\HideSlide1.pptx")

			'Convert the PPT unhidden slides to PDF file format 
			presentation.SaveToPdfOption.ContainHiddenSlides = False
			Dim result As String = "ToPdf.pdf"
			presentation.SaveToFile(result, FileFormat.PDF)

			'View File
			DocumentViewer(result)
		End Sub

		Private Shared Sub DocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace