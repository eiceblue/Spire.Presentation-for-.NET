Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace Properties
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Properties.pptx")

			'Set the DocumentProperty of PPT document
			presentation.DocumentProperty.Application = "Spire.Presentation"
			presentation.DocumentProperty.Author = "E-iceblue"
			presentation.DocumentProperty.Company = "E-iceblue Co., Ltd."
			presentation.DocumentProperty.Keywords = "Demo File"
			presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation."
			presentation.DocumentProperty.Category = "Demo"
			presentation.DocumentProperty.Title = "This is a demo file."
			presentation.DocumentProperty.Subject = "Test"

			'Save the document
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("Output.pptx")
		End Sub
	End Class
End Namespace