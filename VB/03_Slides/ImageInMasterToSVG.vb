Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.IO
Imports System.Text

Namespace ImageInMasterToSVG
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\Data\ImageInMasterToSVG.pptx")

			'Get the master collection
			Dim masterSlide As IMasterSlide = presentation.Masters(0)

			Dim num As Integer = 1
			For i As Integer = 0 To masterSlide.Shapes.Count - 1
				Dim shape As IShape = masterSlide.Shapes(i)
				If TypeOf shape Is SlidePicture Then
					Dim ps As SlidePicture = TryCast(shape, SlidePicture)
					Dim svgByte() As Byte = ps.SaveAsSvgInSlide()
					Dim fs As New FileStream(num & ".svg", FileMode.Create)
					fs.Write(svgByte, 0, svgByte.Length)
					fs.Close()
					num += 1
				End If
			Next i
		End Sub
	End Class
End Namespace
