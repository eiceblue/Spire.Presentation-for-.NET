Imports System.ComponentModel
Imports System.Text
Imports System.Drawing.Printing
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace Print
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Print.pptx")

			'Print
			Dim printerSettings As New PrinterSettings()
			printerSettings.FromPage = 0
			printerSettings.ToPage = presentation.Slides.Count-1
			presentation.Print(printerSettings)
		End Sub
	End Class
End Namespace