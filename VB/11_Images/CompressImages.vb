Imports System
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace CompressImages
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            'Create a PPT document
            Dim presentation As Presentation = New Presentation()
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractImage.pptx")

            For Each slide As ISlide In presentation.Slides
                For Each Spire.Presentation.IShape shape in slide.Shapes
                    If shape is SlidePicture Then
                        Dim ps As SlidePicture = shape as SlidePicture
                        ' Compress image: Remove cropped areas (true) and compress image resolution to 150 DPI 
                        ps.PictureFill.CompressImage(True, 50f)
                    End If
                Next
            Next
            'Save to file
            Dim result As String = "CompressImages_result.pptx"
            presentation.SaveToFile(result, FileFormat.Pptx2013)

            'Launch the PPT file
            System.Diagnostics.Process.Start(result)

        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class
End Namespace