Imports System
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace ToVideo
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            'Create PPT document
            Dim ppt As Presentation = New Presentation()
            'Load the PPT document from disk.
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ToVideo.pptx")

            ' Download ffmpeg before converting PPT to video
            ' Windows: Download ffmpeg-x.x.x-essentials_build.zip file and extract it.
            ' CentOS: Install via command line. Note that a lower ffmpeg version may cause some slide transition effects to fail.
            Dim saveOption As SaveToVideoOption = New SaveToVideoOption(@"..\..\..\..\..\..\Data\ffmpeg-8.1.1-essentials_build\ffmpeg-8.1.1-essentials_build\bin")
            ' Set frame rate
            saveOption.Fps = 10
            ' Set duration for each slide
            saveOption.DurationForEachSlide = 5
            ppt.SaveToVideoOption = saveOption

            Dim result As String = "toVideo.mp4"
            ppt.SaveToFile(result, FileFormat.MP4)
            ppt.Dispose()
        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class
End Namespace