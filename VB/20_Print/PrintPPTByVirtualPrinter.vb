Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace PrintPPTByVirtualPrinter
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

			'Print PowerPoint document to virtual printer (Microsoft XPS Document Writer).
			Dim document As New PresentationPrintDocument(presentation)
			document.PrinterSettings.PrinterName = "Microsoft XPS Document Writer"

			presentation.Print(document)
		End Sub
	End Class
End Namespace