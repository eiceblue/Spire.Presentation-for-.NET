Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.Drawing.Printing

Namespace SilentlyPrintPPTByDefaultPrinter
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

			'Print the PowerPoint document to default printer.
			Dim document As New PresentationPrintDocument(presentation)
			document.PrintController = New StandardPrintController()

			presentation.Print(document)
		End Sub
	End Class
End Namespace