Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace InsertEMFInPPT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\BlankSample_N.pptx")

			'EMF file path
			Dim ImageFile As String = "..\..\..\..\..\..\Data\InsertEMF.emf"

            'Define image size
            Dim img As Image
            img = img.FromFile(ImageFile)

            Dim width As Single=img.Width/1.5f
			Dim height As Single=img.Height/1.5f
			Dim rect As New RectangleF(100, 100, width,height)

			'Append the EMF in slide
			Dim image As IEmbedImage = presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			image.Line.FillType = FillFormatType.None

			'Save the document
			Dim result As String = "InsertEMFInPPT_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace