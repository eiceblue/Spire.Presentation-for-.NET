Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO

Namespace SetPropertiesForTemplate
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'String for .pptx file 
			Dim pptxResult As String = "Output.pptx"

			'String for .odp file 
			Dim odpResult As String = "Output.odp"

			'String for .ppt file 
			Dim pptResult As String = "Output.ppt"

			'Create the .pptx template
			SetPropertiesForTemplate(pptxResult, FileFormat.Pptx2013)

			'Create the .odp template
			SetPropertiesForTemplate(odpResult, FileFormat.ODP)

			'Create the .ppt template
			SetPropertiesForTemplate(pptResult, FileFormat.PPT)

			'Launching the .pptx file.
			Viewer(pptxResult)
		End Sub
		Private Shared Sub SetPropertiesForTemplate(ByVal filePath As String, ByVal fileFormat As FileFormat)
			'Create a document
			Dim presentation As New Presentation()

			'Set the DocumentProperty 
			presentation.DocumentProperty.Application = "Spire.Presentation"
			presentation.DocumentProperty.Author = "E-iceblue"
			presentation.DocumentProperty.Company = "E-iceblue Co., Ltd."
			presentation.DocumentProperty.Keywords = "Demo File"
			presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation."
			presentation.DocumentProperty.Category = "Demo"
			presentation.DocumentProperty.Title = "This is a demo file."
			presentation.DocumentProperty.Subject = "Test"

			'Save to template file
			presentation.SaveToFile(filePath, fileFormat)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace