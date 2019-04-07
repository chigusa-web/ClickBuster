Public Class CEnemy

    ' クラス内変数
    Private _pbImgEnemy As PictureBox  ' 敵の画像
    Private _enemyOwner As Object      ' 敵が出現するオーナー
    Private _rnd As System.Random      ' ランダム変数
    Private _intTokuten As Integer     ' 敵を倒した時の得点
    Private _intMvwd As Integer        ' 敵が動く際の幅
    Private _intMvud As Integer        ' 上下左右に動く
    Private _intMovePattern As Integer ' 敵のパターン

    ' クリックイベント
    Public Event OnClick(ByVal sender As System.Object, ByVal e As EventArgs)

    ' コンストラクタ
    ' owner : 敵が出現するオーナー
    ' rnd   : ランダム変数
    Public Sub New(ByRef owner As Object, ByRef rand As System.Random)

        ' 敵の画像を表示する変数を初期化
        _pbImgEnemy = New PictureBox
        ' オーナーの取得
        _enemyOwner = owner
        ' ランダム変数
        _rnd = rand
        ' 敵が動く際の幅の初期化
        _intMvwd = 10
        ' 上下左右に動く
        _intMvud = 0

        ' 敵画像がクリックされたイベント
        AddHandler _pbImgEnemy.Click, AddressOf DoClick
        ' ランダムな場所に出現
        _RandPlace()
        ' オーナーに画像を貼り付け
        _enemyOwner.Controls.Add(_pbImgEnemy)

    End Sub

    ' デストラクタ
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    ' Clickイベントを発生させる
    Public Sub DoClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RaiseEvent OnClick(Me, e)
    End Sub

    ' 敵画像をランダムな場所に作成
    Private Sub _RandPlace()

        ' 出現場所はオーナーの範囲内にする
        ' 高さ
        Dim intTop As Integer = _rnd.Next(_enemyOwner.Size.Height - _pbImgEnemy.Size.Height)
        _pbImgEnemy.Top = intTop
        ' 横
        Dim intLeft As Integer = _rnd.Next(_enemyOwner.Size.Width - _pbImgEnemy.Size.Width)
        _pbImgEnemy.Left = intLeft

        ' 敵のパターン(画像と得点)をランダムに指定する
        _intMovePattern = _rnd.Next(1, 4)
        _pbImgEnemy.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Select Case _intMovePattern
            Case 1
                '画像読み込み
                _pbImgEnemy.Image = System.Drawing.Image.FromFile("enemy1.gif")
                '得点
                _intTokuten = 500
            Case 2
                '画像読み込み
                _pbImgEnemy.Image = System.Drawing.Image.FromFile("enemy2.gif")
                '得点
                _intTokuten = 1000
            Case 3
                '画像読み込み
                _pbImgEnemy.Image = System.Drawing.Image.FromFile("enemy3.gif")
                '得点
                _intTokuten = 3000
        End Select

    End Sub

    ' 動きを見せる(オーナーの中での動き)
    Public Sub MvEnm()
        Select Case _intMovePattern
            Case 1

                '--------------------
                ' enemy1
                ' 完全ランダムな動き
                '--------------------
                ' 移動先をランダムに決めます
                Dim intMv As Integer = _rnd.Next(1, 5)
                ' 一歩移動(オーナーの範囲をはみ出しそうになったら現状維持)
                Select Case intMv
                    Case 1 '下
                        If (_pbImgEnemy.Top + _pbImgEnemy.Height + _intMvwd) < (_enemyOwner.height) Then
                            _pbImgEnemy.Top += _intMvwd
                        End If
                    Case 2 '上
                        If (_pbImgEnemy.Top - _intMvwd) > 0 Then
                            _pbImgEnemy.Top -= _intMvwd
                        End If
                    Case 3 '右
                        If (_pbImgEnemy.Left + _pbImgEnemy.Width + _intMvwd) < (_enemyOwner.Width) Then
                            _pbImgEnemy.Left += _intMvwd
                        End If
                    Case 4 '左
                        If (_pbImgEnemy.Left - _intMvwd) > 0 Then
                            _pbImgEnemy.Left -= _intMvwd
                        End If
                End Select


            Case 2

                '--------------------------
                ' enemy2
                ' 上下左右にぶつかるまで動く
                '--------------------------
                If _intMvud = 0 Then
                    ' 初回時だけは上下左右どちらに進むか決める
                    _intMvud = _rnd.Next(1, 5)
                End If

                ' 上下左右にぶつかるまで移動する。
                ' ぶつかったらまた上下左右にランダムに移動する。
                Select Case _intMvud
                    Case 1 '下 
                        If (_pbImgEnemy.Top + _pbImgEnemy.Height + _intMvwd) < (_enemyOwner.height) Then
                            _pbImgEnemy.Top += _intMvwd
                        Else
                            _intMvud = _rnd.Next(1, 5)
                        End If
                    Case 2 '上
                        If (_pbImgEnemy.Top - _intMvwd) > 0 Then
                            _pbImgEnemy.Top -= _intMvwd
                        Else
                            _intMvud = _rnd.Next(1, 5)
                        End If
                    Case 3 '右
                        If (_pbImgEnemy.Left + _pbImgEnemy.Width + _intMvwd) < (_enemyOwner.Width) Then
                            _pbImgEnemy.Left += _intMvwd
                        Else
                            _intMvud = _rnd.Next(1, 5)
                        End If
                    Case 4 '左
                        If (_pbImgEnemy.Left - _intMvwd) > 0 Then
                            _pbImgEnemy.Left -= _intMvwd
                        Else
                            _intMvud = _rnd.Next(1, 5)
                        End If
                End Select

            Case 3

                '--------------------
                ' enemy3
                ' ななめに動く
                '--------------------
                If _intMvud = 0 Then
                    ' 初回時だけは上下左右のななめに進むか決める
                    _intMvud = _rnd.Next(1, 5)
                End If

                ' ななめにぶつかるまで移動する。
                ' ぶつかったら跳ね返されるように移動する。
                Select Case _intMvud
                    Case 1 '斜め右下 
                        If (((_pbImgEnemy.Top + _pbImgEnemy.Height + _intMvwd) < (_enemyOwner.height)) And
                            ((_pbImgEnemy.Left + _pbImgEnemy.Width + _intMvwd) < (_enemyOwner.Width))) Then
                            _pbImgEnemy.Top += _intMvwd
                            _pbImgEnemy.Left += _intMvwd
                        Else
                            '次は跳ね返った位置
                            If ((_pbImgEnemy.Top + _pbImgEnemy.Height + _intMvwd) >= (_enemyOwner.height)) Then
                                _intMvud = 2 ' 斜め右上
                            ElseIf ((_pbImgEnemy.Left + _pbImgEnemy.Width + _intMvwd) >= (_enemyOwner.Width)) Then
                                _intMvud = 3 ' 斜め左下
                            End If
                        End If
                    Case 2 '斜め右上
                        If (((_pbImgEnemy.Top - _intMvwd) > 0) And
                            (_pbImgEnemy.Left + _pbImgEnemy.Width + _intMvwd) < (_enemyOwner.Width)) Then
                            _pbImgEnemy.Top -= _intMvwd
                            _pbImgEnemy.Left += _intMvwd
                        Else
                            '次は跳ね返った位置
                            If ((_pbImgEnemy.Top - _intMvwd) <= 0) Then
                                _intMvud = 1 ' 斜め右下
                            ElseIf (_pbImgEnemy.Left + _pbImgEnemy.Width + _intMvwd) >= (_enemyOwner.Width) Then
                                _intMvud = 4 ' 斜め左上
                            End If
                        End If
                    Case 3 '斜め左下
                        If ((_pbImgEnemy.Left - _intMvwd) > 0) And
                            ((_pbImgEnemy.Top + _pbImgEnemy.Height + _intMvwd) < (_enemyOwner.height)) Then
                            _pbImgEnemy.Top += _intMvwd
                            _pbImgEnemy.Left -= _intMvwd
                        Else
                            '次は跳ね返った位置
                            If (_pbImgEnemy.Left - _intMvwd) <= 0 Then
                                _intMvud = 1 ' 斜め右下
                            ElseIf ((_pbImgEnemy.Top + _pbImgEnemy.Height + _intMvwd) >= (_enemyOwner.height)) Then
                                _intMvud = 4 ' 斜め左上
                            End If
                        End If
                    Case 4 '斜め左上
                        If ((_pbImgEnemy.Left - _intMvwd) > 0) And
                            ((_pbImgEnemy.Top - _intMvwd) > 0) Then
                            _pbImgEnemy.Top -= _intMvwd
                            _pbImgEnemy.Left -= _intMvwd
                        Else
                            '次は跳ね返った位置
                            If (_pbImgEnemy.Left - _intMvwd) <= 0 Then
                                _intMvud = 2 ' 斜め右上
                            ElseIf (_pbImgEnemy.Top - _intMvwd) <= 0 Then
                                _intMvud = 3 ' 斜め左下
                            End If
                        End If
                End Select

        End Select
    End Sub

    ' 敵が倒されたイベント
    Public Sub EnemyDown()
        ' 違う場所に出現する
        _RandPlace()
    End Sub

    ' 敵の得点のゲット
    Public Function GetTokuten() As Integer
        GetTokuten = _intTokuten
    End Function

End Class
