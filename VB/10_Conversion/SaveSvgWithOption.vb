Imports System.IO
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SaveSvgWithOption
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim inputFile As String = "..\..\..\..\..\..\Data\SaveSvgWithOption.pptx"

			' Create Presentation object and load the file
			Dim ppt As New Presentation()
			ppt.LoadFromFile(inputFile)

			' Save the underline as decoration when converting to Svg
			ppt.SaveToSvgOption.SaveUnderlineAsDecoration = True

			' Save to Svg
			Dim svgByte() As Byte = ppt.Slides(0).Shapes(0).SaveAsSvgInSlide()
			Dim fs As New FileStream("SaveSvgWithOption" & "1.svg", FileMode.Create)
			fs.Write(svgByte, 0, svgByte.Length)
			fs.Close()

			'Dispose
			ppt.Dispose()

		End Sub

	End Class
End Namespace