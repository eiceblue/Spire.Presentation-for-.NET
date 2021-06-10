Imports System.ComponentModel
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports Spire.Presentation

Namespace RemoveAllDigitalSignatures
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a ppt document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\RemoveAllDigitalSignatures.pptx")

			'Remove all digital signatures
			If ppt.IsDigitallySigned = True Then
				ppt.RemoveAllDigitalSignatures()
			End If
			'Save the document
			Dim output As String = "RemoveAllDigitalSignatures_result.pptx"
			ppt.SaveToFile(output, FileFormat.Pptx2010)
			'Launch the file
			OutputViewer(output)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace