Namespace AddComment
	Partial Public Class Form1
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(Form1))
			Me.btnRun = New Button()
			Me.lblDescription = New Label()
			Me.pbLogo = New PictureBox()
			CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.SuspendLayout()
			' 
			' btnRun
			' 
			Me.btnRun.ImeMode = Windows.Forms.ImeMode.NoControl
			Me.btnRun.Location = New Point(361, 98)
			Me.btnRun.Name = "btnRun"
			Me.btnRun.Size = New Size(75, 23)
			Me.btnRun.TabIndex = 92
			Me.btnRun.Text = "Run"
			Me.btnRun.UseVisualStyleBackColor = True
'			Me.btnRun.Click += New System.EventHandler(Me.btnRun_Click)
			' 
			' lblDescription
			' 
			Me.lblDescription.Font = New Font("Verdana", 8.25F)
			Me.lblDescription.ImeMode = Windows.Forms.ImeMode.NoControl
			Me.lblDescription.Location = New Point(75, 16)
			Me.lblDescription.Name = "lblDescription"
			Me.lblDescription.Size = New Size(358, 79)
			Me.lblDescription.TabIndex = 91
			Me.lblDescription.Text = resources.GetString("lblDescription.Text")
			' 
			' pbLogo
			' 
			Me.pbLogo.Image = (CType(resources.GetObject("pbLogo.Image"), Image))
			Me.pbLogo.ImeMode = Windows.Forms.ImeMode.NoControl
			Me.pbLogo.Location = New Point(16, 16)
			Me.pbLogo.Name = "pbLogo"
			Me.pbLogo.Size = New Size(56, 48)
			Me.pbLogo.SizeMode = Windows.Forms.PictureBoxSizeMode.AutoSize
			Me.pbLogo.TabIndex = 90
			Me.pbLogo.TabStop = False
			' 
			' Form1
			' 
			Me.AutoScaleDimensions = New SizeF(6F, 12F)
			Me.AutoScaleMode = Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New Size(452, 136)
			Me.Controls.Add(Me.btnRun)
			Me.Controls.Add(Me.lblDescription)
			Me.Controls.Add(Me.pbLogo)
			Me.Name = "Form1"
			Me.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
			Me.Text = "Add Comment"
			CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub

		#End Region

		Private WithEvents btnRun As Button
		Private lblDescription As Label
		Private pbLogo As PictureBox
	End Class
End Namespace

