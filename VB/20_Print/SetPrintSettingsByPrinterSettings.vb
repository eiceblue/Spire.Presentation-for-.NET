Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.Drawing.Printing

Namespace SetPrintSettingsByPrinterSettings
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

			'Use PrinterSettings object to print presentation slides.
			Dim ps As New PrinterSettings()
			ps.PrintRange = PrintRange.AllPages
			ps.PrintToFile = True
			Dim result As String = "Result-SetPrintSettingsByPrinterSettingsObject.xps"
			ps.PrintFileName = (result)

			'Print the slide with frame.
			presentation.SlideFrameForPrint = True

			'Print the slide with Grayscale.
			presentation.GrayLevelForPrint = True

			'Print 4 slides horizontal.
			presentation.SlideCountPerPageForPrint = PageSlideCount.Four
			presentation.OrderForPrint = Order.Horizontal

			'Only select some slides to print.
			'presentation.SelectSlidesForPrint("1", "3");

			'Print the document.
			presentation.Print(ps)

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