Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO

Namespace GetBuiltinProperties
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\GetProperties.pptx")

			'Get the builtin properties 
			Dim application As String = presentation.DocumentProperty.Application
			Dim author As String = presentation.DocumentProperty.Author
			Dim company As String = presentation.DocumentProperty.Company
			Dim keywords As String = presentation.DocumentProperty.Keywords
			Dim comments As String = presentation.DocumentProperty.Comments
			Dim category As String = presentation.DocumentProperty.Category
			Dim title As String = presentation.DocumentProperty.Title
			Dim subject As String = presentation.DocumentProperty.Subject

			'Create StringBuilder to save 
			Dim content As New StringBuilder()
			content.AppendLine("DocumentProperty.Application: " & application)
			content.AppendLine("DocumentProperty.Author: " & author)
			content.AppendLine("DocumentProperty.Company " & company)
			content.AppendLine("DocumentProperty.Keywords: " & keywords)
			content.AppendLine("DocumentProperty.Comments: " & comments)
			content.AppendLine("DocumentProperty.Category: " & category)
			content.AppendLine("DocumentProperty.Title: " & title)
			content.AppendLine("DocumentProperty.Subject: " & subject)

			'String for .txt file 
			Dim result As String = "GetBuiltinProperties_Output.txt"

			'Save them to a txt file
			File.WriteAllText(result, content.ToString())

			'Launching the result file.
			Viewer(result)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace