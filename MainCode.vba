Option Explicit

Private Sub CB_Close_Click()
    Unload 検査履歴検索画面
    メインメニュー.Show
End Sub

Private Sub CommandButton1_Click()
    '編集可能リストに追加する動作
    Dim iLoop As Long
    
    ListBox4.AddItem ListBox1.List(ListBox1.ListIndex, 0)
    
    
    '///////////////////////////////////////////
    'リストを全項目巡回する方法
    'For iLoop = 1 To ListBox4.ListCount
    
    'Next
    '////////////////////////////////////////////
    
    iLoop = 1
    
    'ListBox1.List(ListBox1.ListIndex,0))
End Sub

Private Sub CommandButton2_Click()
    ListBox4.RemoveItem ListBox4.ListIndex
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '候補リストから選択リストに追加するコマンド（ボタンクリック時）
    ListBox4.AddItem ListBox1.List(ListBox1.ListIndex, 0)
End Sub

Private Sub ListBox4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '選択リストの不要項目をダブルクリックしたら削除するコマンド（ボタンクリック時）
    ListBox4.RemoveItem ListBox4.ListIndex
End Sub

Private Sub UserForm_Initialize()
    '////////////////////////検査履歴保管場所のパス/////////////////////////////////
    Const sPAth As String = "C:\Users\ibuki_k\OneDrive\Excel"
    '////////////////////////////////////////////////////////////////////

    Dim sBuff As String
    Dim FileList() As String
    Dim iLoopInit As Long
    Dim iAssertNum As Long
    'Initialize
    iLoopInit = 0
    iAssertNum = 1000

    sBuff = Dir(sPAth & "\*.xls*")
    '候補リストを検索し、リストに格納する。※安全対策として1000ファイル以上ある場合はストップする
        '上限数はiAssertNumにて調整可能
    Do While (sBuff <> "") Or (iLoopInit > iAssertNum)
        iLoopInit = iLoopInit + 1
        ReDim Preserve FileList(1 To iLoopInit)
        FileList(iLoopInit) = sBuff
        sBuff = Dir()
    Loop

    '安全上限まで至った場合の注意メッセージ表示
    If iLoopInit > 1000 Then
        MsgBox ("編集候補ファイル数が上限" & iAssertNum & "個に達しました。上限数もしくはパスを確認してください")
    End If
    iLoopInit = 0
    'リストボックスに取得した配列をセットする
    ListBox1.List = FileList
End Sub

Sub MainSearchProcess()
    Dim iLoop As Long
    Dim iInnerNum As Long
    Dim sFileListForRead() As String
    iInnerNum = 0
 
    '検索条件で指定されている条件をIDで管理
    Dim iSrchCriteria() As Long
    '①対象リストが選ばれているかを確認 = 対象のブックのリストを読み込む→1個以上の時のみ処理を行わせる
    If ListBox4.ListCount <> 0 Then
        sFileListForRead = ListBox4.List
    Else
        MsgBox ("対象のファイルが選択されていません。1つ以上選択してください")
        Exit Sub
    End If
    '②検索条件が入力されているかを確認→なしなら対象リストの中身全部を表示するようにする
    For iLoop = 1 To 11
        If Controls("TextBox" & iLoop).Value <> "" Then
            iInnerNum = iInnerNum + 1
            ReDim Preserve iSrchCriteria(1 To iInnerNum)
            iSrchCriteria(iInnerNum) = iLoop
        End If
    Next iLoop
    
    If iInnerNum = 0 Then '検索条件未入力
        '//////////////////////////////////////////////////////////////////////////////////
        '//検索条件なし＝全件表示 を実装(予定)
        '//////////////////////////////////////////////////////////////////////////////////
        Else
        '検索条件入力有…入力内容のチェック機構を整備すべき(フールプルーフを拡充予定）
        '//////////////////////////////////////////////////////////////////////////////////
        '//チェック機構挿入部
        '//////////////////////////////////////////////////////////////////////////////////
            '対象のブックを順に検索していく
        Dim iLoopBook As Long
        For iLoop = 1 To UBound(iSrchCriteria)
                    '何番目の検索キーか?を判定
                    iSrchCriteria (iLoop)
                Next iLoop
        '//////////////////////////////////////////////////////////////////////////////////
        '//検索条件あり＝選択されたブックごとにデータを読込→1条件ごとに順繰り検索していく
        'インプット＝iSrchCruteria()
        '//////////////////////////////////////////////////////////////////////////////////
        End If
    End If

'③検索条件に合致するものを順番に検索していき、合致するものを結果リストに入れる
'④結果リストを結果ウインドウに表示する
    
    
End Sub

Function SrchDate(iSrchCriteria() As Long)
    
End Function
'検索ウィンドウでの挙動
