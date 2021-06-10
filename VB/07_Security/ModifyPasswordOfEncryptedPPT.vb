Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ModifyPasswordOfEncryptedPPT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_4.pptx", "123456")

			'Remove the encryption.
			presentation.RemoveEncryption()

			'Protect the document by setting a new password.
			presentation.Protect("654321")

			Dim result As String = "Result-ModifyPasswordOfEncryptedPptFile.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file.
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace