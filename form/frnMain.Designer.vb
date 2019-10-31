<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.txtPassWord = New System.Windows.Forms.TextBox()
        Me.txtDsn = New System.Windows.Forms.TextBox()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.FDialog = New System.Windows.Forms.OpenFileDialog()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(25, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "パスワード"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(25, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(61, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "データソース"
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(248, 70)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(26, 19)
        Me.btnSelect.TabIndex = 2
        Me.btnSelect.Text = "…"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'txtPassWord
        '
        Me.txtPassWord.Location = New System.Drawing.Point(92, 29)
        Me.txtPassWord.Name = "txtPassWord"
        Me.txtPassWord.Size = New System.Drawing.Size(150, 19)
        Me.txtPassWord.TabIndex = 3
        '
        'txtDsn
        '
        Me.txtDsn.Location = New System.Drawing.Point(92, 70)
        Me.txtDsn.Name = "txtDsn"
        Me.txtDsn.Size = New System.Drawing.Size(150, 19)
        Me.txtDsn.TabIndex = 4
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(104, 114)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(82, 29)
        Me.btnOk.TabIndex = 5
        Me.btnOk.Text = "OK"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'FDialog
        '
        Me.FDialog.InitialDirectory = "C:\Program Files\Common Files\ODBC\Data Sources\"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 161)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.txtDsn)
        Me.Controls.Add(Me.txtPassWord)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "接続設定"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents txtPassWord As System.Windows.Forms.TextBox
    Friend WithEvents txtDsn As System.Windows.Forms.TextBox
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents FDialog As System.Windows.Forms.OpenFileDialog

End Class
