Imports System.IO
Imports Spire.Presentation

Namespace LoadEncryptedStream
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a Presentation instance
			Dim ppt As New Presentation()

			'Load PowerPoint file from stream
			Dim from_stream As FileStream = File.OpenRead("..\..\..\..\..\..\Data\\OpenEncryptedPPT.pptx")

			' The password
			Dim password As String = "123456"

			' Load the encrypted stream with the provided password
			ppt.LoadFromStream(from_stream, FileFormat.Auto, password)

			' Save the decrypted document to disk
			ppt.SaveToFile("output/result.pptx", FileFormat.Pptx2013)

			' Dispose the Presentation object
			ppt.Dispose()
		End Sub
	End Class
End Namespace