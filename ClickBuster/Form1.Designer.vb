<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_Menu
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
        Me.B_Syokyu = New System.Windows.Forms.Button()
        Me.B_Chukyu = New System.Windows.Forms.Button()
        Me.B_Jyokyu = New System.Windows.Forms.Button()
        Me.B_Close = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'B_Syokyu
        '
        Me.B_Syokyu.Location = New System.Drawing.Point(12, 12)
        Me.B_Syokyu.Name = "B_Syokyu"
        Me.B_Syokyu.Size = New System.Drawing.Size(260, 74)
        Me.B_Syokyu.TabIndex = 0
        Me.B_Syokyu.Text = "初級"
        Me.B_Syokyu.UseVisualStyleBackColor = True
        '
        'B_Chukyu
        '
        Me.B_Chukyu.Location = New System.Drawing.Point(12, 92)
        Me.B_Chukyu.Name = "B_Chukyu"
        Me.B_Chukyu.Size = New System.Drawing.Size(260, 74)
        Me.B_Chukyu.TabIndex = 1
        Me.B_Chukyu.Text = "中級"
        Me.B_Chukyu.UseVisualStyleBackColor = True
        '
        'B_Jyokyu
        '
        Me.B_Jyokyu.Location = New System.Drawing.Point(12, 172)
        Me.B_Jyokyu.Name = "B_Jyokyu"
        Me.B_Jyokyu.Size = New System.Drawing.Size(260, 74)
        Me.B_Jyokyu.TabIndex = 2
        Me.B_Jyokyu.Text = "上級"
        Me.B_Jyokyu.UseVisualStyleBackColor = True
        '
        'B_Close
        '
        Me.B_Close.Location = New System.Drawing.Point(12, 275)
        Me.B_Close.Name = "B_Close"
        Me.B_Close.Size = New System.Drawing.Size(260, 74)
        Me.B_Close.TabIndex = 3
        Me.B_Close.Text = "終了"
        Me.B_Close.UseVisualStyleBackColor = True
        '
        'F_Menu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 361)
        Me.Controls.Add(Me.B_Close)
        Me.Controls.Add(Me.B_Jyokyu)
        Me.Controls.Add(Me.B_Chukyu)
        Me.Controls.Add(Me.B_Syokyu)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(1, 2, 1, 2)
        Me.MinimizeBox = False
        Me.Name = "F_Menu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ClickBuster (メニュー画面)"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents B_Syokyu As Button
    Friend WithEvents B_Chukyu As Button
    Friend WithEvents B_Jyokyu As Button
    Friend WithEvents B_Close As Button
End Class
