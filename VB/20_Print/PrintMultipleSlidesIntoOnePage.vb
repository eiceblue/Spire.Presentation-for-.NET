Imports System.ComponentModel
Imports System.Text
Imports System.Drawing.Printing
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace PrintMultipleSlidesIntoOnePage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			'Create a PPT document
			Dim ppt As New Presentation()
			'Load the document from disk
			ppt.LoadFromFile("..\..\..\..\..\..\Data\PrintMultipleSlidesIntoOnePage.pptx")
			Dim document As New PresentationPrintDocument(ppt)

			'Set print task name
			document.DocumentName = "print task 1"
			document.PrintOrder = Order.Horizontal
			document.SlideFrameForPrint = True

			'Set Gray level when printing
			document.GrayLevelForPrint = True
			'Set four slides on one page
			document.SlideCountPerPageForPrint = PageSlideCount.Four

			'Set continuous print area
			document.PrinterSettings.PrintRange = PrintRange.SomePages
			document.PrinterSettings.FromPage = 1
			document.PrinterSettings.ToPage = ppt.Slides.Count - 1

			'Set discontinuous print area
			'document.SelectSldiesForPrint("1", "2-4");

			ppt.Print(document)
			ppt.Dispose()
		End Sub
	End Class
End Namespace