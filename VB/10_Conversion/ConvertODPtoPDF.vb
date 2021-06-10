Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ConvertODPtoPDF
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			Dim presentation As New Presentation()

			'Load ODP file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\toPdf.odp",FileFormat.ODP)

			Dim result As String = "ConvertODPtoPDF_result.pdf"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.PDF)

			'Launch the PowerPoint file.
			DocumentViewer(result)
		End Sub

		Private Sub DocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace