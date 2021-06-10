Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports System.Text
Imports System.IO
Imports Spire.Presentation

Namespace ToTIFF
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			'Create PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ToTIFF.pptx")
			Dim images(presentation.Slides.Count - 1) As Image

			'Save PPT to images
			For i As Integer = 0 To presentation.Slides.Count - 1
				images(i) = presentation.Slides(i).SaveAsImage()
			Next i

			'Make TIFF image using images
			JoinTiffImages(images, "ToTIFF.tiff", EncoderValue.CompressionLZW)

			Process.Start("ToTIFF.tiff")
		End Sub

		'Function to get specified ImageCodecInfo
		Private Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
			Dim encoders() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders()
			For j As Integer = 0 To encoders.Length - 1
				If encoders(j).MimeType = mimeType Then
					Return encoders(j)
				End If
			Next j

			Throw New Exception(mimeType & " mime type not found in ImageCodecInfo")
		End Function

		'Function to make TIFF using images
		Public Shared Sub JoinTiffImages(ByVal images() As Image, ByVal outFile As String, ByVal compressEncoder As EncoderValue)
			'Use the save encoder
			Dim enc As System.Drawing.Imaging.Encoder = System.Drawing.Imaging.Encoder.SaveFlag

			Dim ep As New EncoderParameters(2)
			ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.MultiFrame))
			ep.Param(1) = New EncoderParameter(System.Drawing.Imaging.Encoder.Compression, CLng(compressEncoder))

			Dim pages As Image = Nothing
			Dim frame As Integer = 0
			Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")

			For Each img As Image In images
				If frame = 0 Then
					pages = img

					'Save the first frame
					pages.Save(outFile, info, ep)
				Else
					'Save the intermediate frames
					ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.FrameDimensionPage))

					pages.SaveAdd(img, ep)
				End If

				If frame = images.Length - 1 Then
					'Flush and close
					ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.Flush))
					pages.SaveAdd(ep)
				End If

				frame += 1
			Next img
		End Sub

	End Class
End Namespace