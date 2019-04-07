Public Class F_Main

    Public Property Enms As Integer          ' 敵の数
    Public Property EnmTim As Integer        ' 敵の動くスピード(秒数)
    Public Property GmTim As Integer         ' 残り時間(秒数)

    Private _rnd As New System.Random    ' ランダム変数
    Private _enemies As New ArrayList()  ' 複数の敵を格納する変数

    ' 難易度初期化
    Public Sub FrmIni()

        ' フォームのサイズでコントロールを配置する
        P_Enemy.Height = Me.Height - P_Enemy.Top - 45
        P_Enemy.Width = Me.Width - 40

        ' 合計得点
        L_Sum.Text = "0"
        ' 敵の動くスピード
        T_Enemy.Interval = EnmTim
        ' 残り時間の初期化
        PG_Jikan.Maximum = (GmTim / 1000)
        PG_Jikan.Value = PG_Jikan.Maximum
    End Sub

    Private Sub F_Main_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        ' メニュー画面の表示
        F_Menu.Show()
    End Sub

    ' スタートボタンクリック
    Private Sub B_Start_Click(sender As Object, e As EventArgs) Handles B_Start.Click
        ' スタートボタンを使えなくする
        B_Start.Enabled = False

        ' 敵の作成
        Dim i As Integer
        ' 敵を10体生成します
        For i = 0 To Enms - 1
            ' 配列に敵クラスを生成します。その際にオーナーのパネルとランダム変数を渡します。
            _enemies.Add(New CEnemy(Me.P_Enemy, _rnd))
            ' 敵のクリックイベント(倒すイベント)
            AddHandler CType(_enemies(i), CEnemy).OnClick, AddressOf EnmOnClick
            ' ループの中でDoEventsを呼んでおく
            System.Windows.Forms.Application.DoEvents()
        Next

        ' タイマーのスタート
        T_Enemy.Enabled = True
        T_Jikan.Enabled = True
    End Sub

    ' 敵を動かす
    Private Sub T_Enemy_Tick(sender As Object, e As EventArgs) Handles T_Enemy.Tick
        Dim i As Integer
        ' 敵の数だけループする
        For i = 0 To _enemies.Count() - 1
            ' 敵の移動イベントを呼ぶ
            CType(_enemies(i), CEnemy).MvEnm()
            ' ループの中でDoEventsを呼んでおく
            System.Windows.Forms.Application.DoEvents()
        Next
    End Sub

    ' 残り時間を計測
    Private Sub T_Jikan_Tick(sender As Object, e As EventArgs) Handles T_Jikan.Tick
        If (PG_Jikan.Minimum) <= (PG_Jikan.Value - 1) Then
            ' プログレスバーを一つ下げる
            PG_Jikan.Value = PG_Jikan.Value - 1
        Else
            ' 終わり
            PG_Jikan.Value = PG_Jikan.Minimum
            ' タイマーを止める
            T_Enemy.Enabled = False
            T_Jikan.Enabled = False
            ' メッセージを表示
            MsgBox("終了です" & vbCrLf & "あなたの得点は" & L_Sum.Text & "です")
            ' メイン画面を閉じる
            Me.Close()
        End If
    End Sub

    ' 敵クリックイベント
    Public Sub EnmOnClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' 得点の加算(それぞれの敵の得点を取得する。)
        L_Sum.Text = CInt(L_Sum.Text) + CType(sender, CEnemy).GetTokuten
        ' 敵を倒したイベント(再度敵を出現させる)
        Call CType(sender, CEnemy).EnemyDown()
    End Sub

End Class