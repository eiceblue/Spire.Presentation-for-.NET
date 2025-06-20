Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Collections
Imports Spire.Presentation.External.Pdf

Namespace ToPDFA
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ToPDF.pptx")

			'Save the PPT to PDF_A1A
			ppt.SaveToPdfOption.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1A
			Dim result As String = "ToPDF_A1A.pdf"
			ppt.SaveToFile(result, FileFormat.PDF)

			'Save the PPT to PDF_A1B
			ppt.SaveToPdfOption.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B
			result = "ToPDF_A1B.pdf"
			ppt.SaveToFile(result, FileFormat.PDF)

			'Save the PPT to PDF_A2A
			ppt.SaveToPdfOption.PdfConformanceLevel = PdfConformanceLevel.Pdf_A2A
			result = "ToPDF_A2A.pdf"
			ppt.SaveToFile(result, FileFormat.PDF)

			'View the document
			FileViewer(result)

		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace