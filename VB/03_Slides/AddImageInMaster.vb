Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace AddImageInMaster
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AddImageInMaster.pptx")

			'Get the master collection
			Dim master As IMasterSlide = presentation.Masters(0)

			'Append image to slide master
			Dim image As String = "..\..\..\..\..\..\Data\Logo.png"
			Dim rff As New RectangleF(40, 40, 90, 90)
			Dim pic As IEmbedImage = master.Shapes.AppendEmbedImage(ShapeType.Rectangle, image, rff)
			pic.Line.FillFormat.FillType = FillFormatType.None

			'Add new slide to presentation
			presentation.Slides.Append()

			'Save the document
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010)

			'Launch the PPT file
			Process.Start("Output.pptx")
		End Sub
	End Class
End Namespace
