Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO
Imports System.Drawing.Printing

Namespace SpecificPrinterToPrint
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT document from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ChangeSlidePosition.pptx")

			'New PrintSeetings
			Dim printerSettings As New PrinterSettings()

			'Set landscape for page
			printerSettings.DefaultPageSettings.Landscape = True

			'Specific the printer
			printerSettings.PrinterName = "Microsoft XPS Document Writer"

			'Print 
			presentation.Print(printerSettings)
		End Sub
	End Class
End Namespace