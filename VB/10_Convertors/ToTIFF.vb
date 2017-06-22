Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Text
Imports System.Windows.Forms
Imports System.IO

Public Class Form1

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click

        'create PPT document
        Dim presentation As New Presentation()

        'load PPT file from disk
        presentation.LoadFromFile("..\..\..\..\..\..\Data\source.pptx")
        Dim images As Image() = New Image(presentation.Slides.Count - 1) {}

        'save PPT to images
        For i As Integer = 0 To presentation.Slides.Count - 1
            images(i) = presentation.Slides(i).SaveAsImage()
        Next

        'make TIFF image using images
        JoinTiffImages(images, "result.tiff", EncoderValue.CompressionLZW)

        System.Diagnostics.Process.Start("result.tiff")

    End Sub

    'function to get specified ImageCodecInfo
    Private Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
        Dim encoders As ImageCodecInfo() = ImageCodecInfo.GetImageEncoders()
        For j As Integer = 0 To encoders.Length - 1
            If encoders(j).MimeType = mimeType Then
                Return encoders(j)
            End If
        Next

        Throw New Exception(mimeType & Convert.ToString(" mime type not found in ImageCodecInfo"))
    End Function

    'function to make TIFF using images
    Public Shared Sub JoinTiffImages(ByVal images As Image(), ByVal outFile As String, ByVal compressEncoder As EncoderValue)
        'use the save encoder
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

                'save the first frame
                pages.Save(outFile, info, ep)
            Else
                'save the intermediate frames
                ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.FrameDimensionPage))

                pages.SaveAdd(img, ep)
            End If

            If frame = images.Length - 1 Then
                'flush and close.
                ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.Flush))
                pages.SaveAdd(ep)
            End If

            frame += 1
        Next
    End Sub

End Class