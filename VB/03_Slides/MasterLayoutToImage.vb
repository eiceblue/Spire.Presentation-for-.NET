Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace MasterLayoutToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a PPT document and load the file 
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\CloneMaster2.pptx")

			' Iterate the masters
			For i As Integer = 0 To ppt.Masters(0).Layouts.Count - 1
				' Save layouts as images
				Dim image As Image = ppt.Masters(0).Layouts(i).SaveAsImage()
				Dim fileName As String = String.Format("{0}.png", i)
				image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png)

				'////////////////Use the following code for netstandard dlls/////////////////////////
'                
'                using (var images = ppt.Masters[0].Layouts[i].SaveAsImage())
'                {
'                    String filename = String.Format("MasterLayoutToImage_{0}.png", i);
'                    FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);                    
'                    images.Seek(0, SeekOrigin.Begin);
'                    images.CopyTo(fileStream);
'                    fileStream.Flush();
'                    fileStream.Close();
'                }
'                

			Next i

			ppt.Dispose()
		End Sub
	End Class
End Namespace