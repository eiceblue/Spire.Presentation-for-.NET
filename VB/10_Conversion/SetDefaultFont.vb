Imports System
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace SetDefaultFont
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            'Specify the input and output files
            Dim inputFile As String = @"..\..\..\..\..\..\Data\ToPDF_DefaultFont.pptx"
            Dim outputFile As String = "SetDefaultFont.pdf"

            'Set the default fonts used for the conversion
            Presentation.SetDefaultLatinFontName("Arial")
            Presentation.SetDefaultEastAsianFontName("微软雅黑")

            'Convert the PPT document to PDF
            Dim ppt As Presentation = New Presentation()
            ppt.LoadFromFile(inputFile)
            ppt.SaveToFile(outputFile, FileFormat.PDF)

            'Reset the default font settings
            Presentation.ResetDefaultEastAsianFontName()
            Presentation.ResetDefaultLatinFontName()

            'Launch the PDF file
            System.Diagnostics.Process.Start(outputFile)
        End Sub
    End Class
End Namespace
