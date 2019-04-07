<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_Main
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
        Me.components = New System.ComponentModel.Container()
        Me.B_Start = New System.Windows.Forms.Button()
        Me.L_Jikan = New System.Windows.Forms.Label()
        Me.L_Tokuten = New System.Windows.Forms.Label()
        Me.L_Sum = New System.Windows.Forms.Label()
        Me.PG_Jikan = New System.Windows.Forms.ProgressBar()
        Me.P_Enemy = New System.Windows.Forms.Panel()
        Me.T_Enemy = New System.Windows.Forms.Timer(Me.components)
        Me.T_Jikan = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'B_Start
        '
        Me.B_Start.Location = New System.Drawing.Point(12, 12)
        Me.B_Start.Name = "B_Start"
        Me.B_Start.Size = New System.Drawing.Size(80, 49)
        Me.B_Start.TabIndex = 0
        Me.B_Start.Text = "スタート"
        Me.B_Start.UseVisualStyleBackColor = True
        '
        'L_Jikan
        '
        Me.L_Jikan.AutoSize = True
        Me.L_Jikan.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.L_Jikan.Location = New System.Drawing.Point(98, 12)
        Me.L_Jikan.Name = "L_Jikan"
        Me.L_Jikan.Size = New System.Drawing.Size(66, 16)
        Me.L_Jikan.TabIndex = 1
        Me.L_Jikan.Text = "残り時間"
        '
        'L_Tokuten
        '
        Me.L_Tokuten.AutoSize = True
        Me.L_Tokuten.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.L_Tokuten.Location = New System.Drawing.Point(249, 12)
        Me.L_Tokuten.Name = "L_Tokuten"
        Me.L_Tokuten.Size = New System.Drawing.Size(40, 16)
        Me.L_Tokuten.TabIndex = 2
        Me.L_Tokuten.Text = "得点"
        '
        'L_Sum
        '
        Me.L_Sum.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.L_Sum.Location = New System.Drawing.Point(249, 45)
        Me.L_Sum.Name = "L_Sum"
        Me.L_Sum.Size = New System.Drawing.Size(123, 16)
        Me.L_Sum.TabIndex = 3
        Me.L_Sum.Text = "0"
        Me.L_Sum.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'PG_Jikan
        '
        Me.PG_Jikan.Location = New System.Drawing.Point(101, 35)
        Me.PG_Jikan.Maximum = 10
        Me.PG_Jikan.Name = "PG_Jikan"
        Me.PG_Jikan.Size = New System.Drawing.Size(142, 26)
        Me.PG_Jikan.TabIndex = 4
        Me.PG_Jikan.Value = 10
        '
        'P_Enemy
        '
        Me.P_Enemy.BackColor = System.Drawing.Color.White
        Me.P_Enemy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.P_Enemy.Location = New System.Drawing.Point(12, 67)
        Me.P_Enemy.Name = "P_Enemy"
        Me.P_Enemy.Size = New System.Drawing.Size(360, 402)
        Me.P_Enemy.TabIndex = 5
        '
        'T_Enemy
        '
        Me.T_Enemy.Interval = 1000
        '
        'T_Jikan
        '
        Me.T_Jikan.Interval = 1000
        '
        'F_Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 481)
        Me.Controls.Add(Me.P_Enemy)
        Me.Controls.Add(Me.PG_Jikan)
        Me.Controls.Add(Me.L_Sum)
        Me.Controls.Add(Me.L_Tokuten)
        Me.Controls.Add(Me.L_Jikan)
        Me.Controls.Add(Me.B_Start)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimizeBox = False
        Me.Name = "F_Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ClickBuster (メイン画面)"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents B_Start As Button
    Friend WithEvents L_Jikan As Label
    Friend WithEvents L_Tokuten As Label
    Friend WithEvents L_Sum As Label
    Friend WithEvents PG_Jikan As ProgressBar
    Friend WithEvents P_Enemy As Panel
    Friend WithEvents T_Enemy As Timer
    Friend WithEvents T_Jikan As Timer
End Class
