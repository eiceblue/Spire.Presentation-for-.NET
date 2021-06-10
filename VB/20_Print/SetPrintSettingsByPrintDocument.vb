Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetPrintSettingsByPrintDocument
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

			'Use PrintDocument object to print presentation slides.
			Dim document As New PresentationPrintDocument(presentation)

			'Print document to virtual printer.
			document.PrinterSettings.PrinterName = "Microsoft XPS Document Writer"

			'Print the slide with frame.
			presentation.SlideFrameForPrint = True

			'Print 4 slides horizontal.
			presentation.SlideCountPerPageForPrint = PageSlideCount.Four
			presentation.OrderForPrint = Order.Horizontal

			'Print the slide with Grayscale.
			presentation.GrayLevelForPrint = True

			'Set the print document name.          
			document.DocumentName = "Template_Ppt_6.pptx"

			document.PrinterSettings.PrintToFile = True
			Dim result As String = "Result-SetPrintSettingsByPrintDocumentObject.xps"
			document.PrinterSettings.PrintFileName = (result)

			presentation.Print(document)

			'Launch the file.
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace