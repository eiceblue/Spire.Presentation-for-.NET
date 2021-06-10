Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.Drawing.Printing

Namespace PrintSpecifiedRangeOfPPTPages
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_6.pptx")

			Dim document As New PresentationPrintDocument(presentation)

			'Set the document name to display while printing the document. 
			document.DocumentName = "Template_Ppt_6.pptx"

			'Choose to print some pages from the PowerPoint document.
			document.PrinterSettings.PrintRange = PrintRange.SomePages
			document.PrinterSettings.FromPage = 2
			document.PrinterSettings.ToPage = 3

			'Set the number of copies of the document to print.
			document.PrinterSettings.Copies = 2

			presentation.Print(document)
		End Sub
	End Class
End Namespace