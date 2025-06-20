Imports Spire.Presentation


Namespace EmbedSvg
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			Spire.Presentation.License.LicenseProvider.SetLicenseKey("p64ePqlI7i7yWMfAAQDpUfp+nvC6HIUNbmVmuR9vr54bk4UGKOLXeOPNQqBlBinudyeNG0YvYSRaIXjddges/6BpzKGGg2qBjuFAwDIKXGgi4pxJPhQMz+iZl+j+CpQlrSX2MQd6lLfTCe4VptqCiaTsO/XhKJTqPyu/jQ0Ybt35OEoTN05XXtptZU0bV8HdCcvYv9w7U5XXbXtz9OCD0sKKTvrIcuoPGYBj/LBNKrkh+7s9zpEUvn5gmEuUtKdZhqd7qt47qLBPFDagADsU1C/EfIrG+bOtrtqYdDPCG43jaeBM7hTLsGzwqKSmt1QA7ZxytXQVkZ5R7MPXBwOVSjojT7xB2lxNi+yZ8Y9bQAj1SgL7mdpqfMCveoM3IdA3o7ADlvlqvtFQHEhvZU6DrYsGyRxaxpBLuFyC105EQ3VSdzbf0fmHRmfetCXJeN/09smcmtAm+e4Cru+My1HN0ubGfZKA9pVo6R6+lFBPaOL4MqamOM96ZlIfQ32nhWoqjvsU1P1WQbFhJ0WTwaF/Ak9D6oM/31/iAlXTPP47Me43MF9YYQEyZPVuCqk8gz6LEf3JbpeJpE5QGVvirYJ2hsvVMvXJK3tbCAQ5XQ7Gf7fkwL22nrpBO/9SqudWHEKnLae17buO92V5mtZR0SdcLw7nX9rfEVHW9f2bSIeGtMTFxJTRgPn2cwwm/WRg9EkjVQ4AF45wer5AiBRBaWGPpo++xRZAnLe3Nl+t6ZBFaGW/7MBvEdRyTNaRl1sbcWKztAnLgxGd0xlh0pu4MdntdBv9smoCLF8w70bGtUhnrVYZGG/KSPsVUOkxtGHHSflEnhkplSlrcEJ/nYOoSRQfgVipjUDGfcOsoWrSv5eVTe1N2FAna3LNbMgm6bB0zB+WmQ2X9LxmJ01Ux8u8J2aDewOJoeR/BL4s/OE0PygeSQSH3rLi/j5Yl5s0315MM9avLdeHacWHjbO47ClR+K9z7FYTaPPqx9WMYXMNYYon4Nbd4EUAcC/w1O2dgeL26WOyTWsJOs3f6dmEJrsCBXQsfZet5tT7lv7roonL1e7+/HsgDHJhStsguhRJsJ/s85no0zUNEnFS0fG68DHVa3TRG4Nx1UYYin0QMhs56IpBfyIwpUmWXkRHjfaZC7Lb1KhLt8nk2SJ/H7imX1jciSQ1zS7JggLXGA66ZSIU89l+dmZkzmetYo0KB5oBYQS8nioP/1IAhzKIFMqBJWEH4KMag4BEheBWZhluOIA+XSXT/3s64G18MwsLqsrj1s/MoctCKZDLNAC+fONjzV5B1lhXntRwubKO1ims2I2MaxMkKdcgMWUbndV4JTFt5IhvtVFph9BiSlyoYUp9oJdwg1xeIBx60mRYEztFPkbkGMUaI2uMenmWTfWQ3ddMTm7KNw1kx9uTWbl0G845kq/x0ZDoViS0SaBY21OP+PiFeBAJqluWYs8HAE9V1f4dER80tQlUWkkAwmtasylnZx5KF2zCiyMEb3BvDICyuWMOnSmy1YRR9UHciNeFD4lWffTccJp4WYNsI81OZ8/+9cs7Mz/U46manclamKx2Q/enWzfL9sCdn/EURd+mx7CWcN30c3+035/a6nEnmo+5KIcNlLMPDEnbKyUSEwBn")

			Dim inputFile As String = "..\..\..\..\..\..\Data\charthtml.svg"

			' Create Presentation object
			Dim presentation As New Presentation()

			' Embed svg in presentation shape
			presentation.Slides(0).Shapes.AddFromSVG(inputFile, New RectangleF(40, 40, 200, 200))

			' Save the file
			presentation.SaveToFile("EmbedSvg_output2.pptx",FileFormat.Pptx2019)

			'Dispose
			presentation.Dispose()

			'System.Diagnostics.Process.Start("EmbedSvg_output.pptx");
			Me.Close()
		End Sub
	End Class
End Namespace