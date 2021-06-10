Imports System.ComponentModel
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports Spire.Presentation

Namespace AddDigitalSignature
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a ppt document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\AddDigitalSignature.pptx")

			'Load the certificate
			Dim x509 As New X509Certificate2("..\..\..\..\..\..\Data\gary.pfx", "e-iceblue")

			'Add a digital signature
			ppt.AddDigitalSignature(x509, "111", Date.Now)

			'Save the document
			ppt.SaveToFile("AddDigitalSignature_result.pptx", FileFormat.Pptx2010)
			Process.Start("AddDigitalSignature_result.pptx")
		End Sub
	End Class
End Namespace