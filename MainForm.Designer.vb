<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form, bileşen listesini temizlemeyi bırakmayı geçersiz kılar.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form Tasarımcısı tarafından gerektirilir
    Private components As System.ComponentModel.IContainer

    'NOT: Aşağıdaki yordam Windows Form Tasarımcısı için gereklidir
    'Windows Form Tasarımcısı kullanılarak değiştirilebilir.  
    'Kod düzenleyicisini kullanarak değiştirmeyin.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.btnSelectFile = New System.Windows.Forms.Button()
        Me.Aciklama = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'btnSelectFile
        '
        Me.btnSelectFile.Location = New System.Drawing.Point(11, 11)
        Me.btnSelectFile.Name = "btnSelectFile"
        Me.btnSelectFile.Size = New System.Drawing.Size(280, 36)
        Me.btnSelectFile.TabIndex = 0
        Me.btnSelectFile.Text = "Temizlenecek XLSX dosyasını seçiniz..."
        Me.btnSelectFile.UseVisualStyleBackColor = True
        '
        'Aciklama
        '
        Me.Aciklama.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Aciklama.AutoSize = True
        Me.Aciklama.Location = New System.Drawing.Point(12, 50)
        Me.Aciklama.MaximumSize = New System.Drawing.Size(280, 90)
        Me.Aciklama.Name = "Aciklama"
        Me.Aciklama.Size = New System.Drawing.Size(40, 13)
        Me.Aciklama.TabIndex = 1
        Me.Aciklama.Text = "Hazır..."
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(11, 139)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(280, 10)
        Me.ProgressBar1.TabIndex = 2
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(304, 161)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Aciklama)
        Me.Controls.Add(Me.btnSelectFile)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "INC V1.1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnSelectFile As Button
    Friend WithEvents Aciklama As Label
    Friend WithEvents ProgressBar1 As ProgressBar
End Class
