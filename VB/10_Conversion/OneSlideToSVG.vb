Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports System.Text
Imports System.IO
Imports Spire.Presentation

Namespace OneSlideToSVG
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			'Create PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\OneSlideToSVG.pptx")

			'Convert the second slide to SVG
			Dim svgByte() As Byte = presentation.Slides(1).SaveToSVG()
			File.WriteAllBytes("OneSlideToSVG.svg", svgByte)
            'Launch the file
            OutputViewer("OneSlideToSVG.svg")
        End Sub
        Private Sub OutputViewer(ByVal filename As String)
            Try
                Process.Start(filename)
            Catch
            End Try
        End Sub
    End Class
End Namespace