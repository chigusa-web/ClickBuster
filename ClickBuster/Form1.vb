Public Class F_Menu

    ' 初級ボタンクリック
    Private Sub B_Syokyu_Click(sender As Object, e As EventArgs) Handles B_Syokyu.Click
        F_Main.Height = 400    ' フォームの高さ
        F_Main.Width = 400     ' フォームの幅
        F_Main.EnmTim = 1000   ' 敵の動く速度
        F_Main.Enms = 20       ' 敵の数
        F_Main.GmTim = 10000   ' 制限時間
        F_Main.FrmIni()        ' 難易度初期化
        F_Main.Show()
        Me.Hide()
    End Sub

    ' 中級ボタンクリック
    Private Sub B_Chukyu_Click(sender As Object, e As EventArgs) Handles B_Chukyu.Click
        F_Main.Height = 500    ' フォームの高さ
        F_Main.Width = 700     ' フォームの幅
        F_Main.EnmTim = 800    ' 敵の動く速度
        F_Main.Enms = 30       ' 敵の数
        F_Main.GmTim = 15000   ' 制限時間
        F_Main.FrmIni()        ' 難易度初期化
        F_Main.Show()
        Me.Hide()
    End Sub

    ' 上級ボタンクリック
    Private Sub B_Jyokyu_Click(sender As Object, e As EventArgs) Handles B_Jyokyu.Click
        F_Main.Height = 600    ' フォームの高さ
        F_Main.Width = 800     ' フォームの幅
        F_Main.EnmTim = 300    ' 敵の動く速度
        F_Main.Enms = 40       ' 敵の数
        F_Main.GmTim = 20000   ' 制限時間
        F_Main.FrmIni()        ' 難易度初期化
        F_Main.Show()
        Me.Hide()
    End Sub

    ' 終了ボタンクリック
    Private Sub B_Close_Click(sender As Object, e As EventArgs) Handles B_Close.Click
        ' メニュー画面を閉じる
        Me.Close()
    End Sub
End Class
