Attribute VB_Name = "check_required"
'# __author__  = "Hiroshi Ohta"
'# __version__ = "0.01"
'# __date__    = "01 Nov 2019"

Sub CheckRequired()

    Application.ScreenUpdating = False
    
    '# 定数定義
    '# -----------------------------------------------------------------------------------------------------------------
    Dim CATEGORY_LIST As Variant
    Dim SH As Worksheet
    
    
    Dim EXLS_MAX_COLUM As Long
    Dim EXLS_MAX_ROW As Long
    Dim CHECK_DEPS_MAX As Integer
    
    
    '# 変数定義
    '# -----------------------------------------------------------------------------------------------------------------
    Dim def_category As Variant
    Dim column_cnt As Long
    Dim row_cnt As Long
    Dim data_column As Long
    Dim data_row As Long
    Dim check_row As Integer
    Dim def_data As Variant
    
    Dim check_required As Integer
    Dim check_index_row As Integer
    Dim error_cnt As Integer
    Dim required_message As String
    Dim required_key As String
    Dim insert_column As Integer
    Dim check_deps_cnt As Integer
    Dim check_deps_result As Variant
    Dim check_deps_message As String
    
    
    '# 定数設定
    '# -----------------------------------------------------------------------------------------------------------------
    CATEGORY_LIST = Array("hst", "tgrp", "job", "fmt", "mfmt", "snd", "rcv", "trg")
    EXLS_MAX_COLUM = 16384
    EXLS_MAX_ROW = 1048576
    
    CHECK_DEPS_MAX = 100
    
    
    '# 定義チェック
    '# -----------------------------------------------------------------------------------------------------------------
    required_message = ""
    def_category_cnt = 0
    
    '# 前処理
    '# -----------------------------------------------------------------------------------------------------------------
    hist.Select
    Range(Cells(3, 17), Cells(CHECK_DEPS_MAX, 17 + 8)).ClearContents
    
    For Each def_category In CATEGORY_LIST
        def_category_cnt = def_category_cnt + 1
        '# 該当シートの取得
        For Each SH In Worksheets
            If def_category = SH.CodeName Then
                Sheets(SH.Name).Select
                Exit For
            End If
        Next
        
        '# 定義情報の取得
        '# -----------------------------------------------------------------------------------------------------------------
        def_data = GetDefData(def_category)
        
        
        '# 必須項目の確認
        '# -----------------------------------------------------------------------------------------------------------------
        required_key = ""
        check_required = 7
        check_start_index_row = check_required + 2
        
        If def_category = "tgrp" Then
            check_start_index_row = check_required + 3
        ElseIf def_category = "fmt" Or def_category = "mfmt" Then
            check_start_index_row = check_required + 4
        End If
        
        For check_index_column = LBound(def_data, 2) To UBound(def_data, 2) Step 1
            If def_data(check_required, check_index_column) = "○" Then
                error_cnt = 0
                For check_index_row = check_start_index_row To UBound(def_data, 1) Step 1
                    If def_data(check_index_row, check_index_column) = "" Then
                        
                        '# 一つのIDで複数の複数行の定義を作成する場合
                        If check_index_column = 1 And (def_category = "tgrp" Or def_category = "fmt" Or def_category = "mfmt") Then
                            If def_category = "tgrp" Then
                                If def_data(check_index_row, 2) = "" Then
                                    error_cnt = error_cnt + 1
                                End If
                            Else
                                If def_data(check_index_row, 6) = "" Then
                                    error_cnt = error_cnt + 1
                                End If
                            End If
                        Else
                            error_cnt = error_cnt + 1
                        End If
                    End If
                Next
                If error_cnt > 0 Then
                    required_key = required_key & vbCrLf & " - " & def_data(1, check_index_column)
                End If
            End If
        Next
        
        If required_key <> "" Then
            required_message = required_message & vbCrLf & vbCrLf & "シート名：" & SH.Name & required_key
        End If
        
        
        '# 依存項目の確認
        '# -----------------------------------------------------------------------------------------------------------------
        '# tgrp → hst, mfmt → fmt, snd → job tgrp fmt mfmt, rcv → job tgrp, trg → job
        hist.Select
        If def_category <> "hst" And def_category <> "job" And def_category <> "fmt" Then
        
            '# 転送グループの確認
            If def_category = "tgrp" Then
                check_deps_column = 17
                Dim check_deps_data(100) As String
                For check_deps_cnt = 0 To CHECK_DEPS_MAX
                    check_deps_data(check_deps_cnt) = Cells(check_deps_cnt + 3, check_deps_column)
                Next
                
                check_deps_message = ""
                
                For check_index_row = check_start_index_row To UBound(def_data, 1) Step 1
                    check_index_column = 2
                    check_deps_result = Filter(check_deps_data, def_data(check_index_row, check_index_column), True)
                    If UBound(check_deps_result) = -1 Then
                        '# errr
                        check_deps_message = check_deps_message & vbCrLf & "  - " & def_data(check_index_row, check_index_column)
                    End If
                Next
                
                If check_deps_message <> "" Then
                    check_deps_message = "次の" & def_data(1, check_index_column) & "は、詳細ホスト情報に定義されてません。" & check_deps_message
                    MsgBox check_deps_message, vbOKOnly + vbExclamation, "入力エラー"
                End If
                
                Erase check_deps_data
                Erase check_deps_result
                
            End If
        End If
        
        
        
        
        '# 履歴登録シートのチェックへデータ投入
        '# -----------------------------------------------------------------------------------------------------------------
        insert_column = def_category_cnt + 16
        
        insert_row = 3
        For check_index_row = check_start_index_row To UBound(def_data, 1) Step 1
            If def_data(check_index_row, 1) <> "" Then
                Cells(insert_row, insert_column) = def_data(check_index_row, 1)
                insert_row = insert_row + 1
            End If
        Next
        
        
        '# 配列の開放
        Erase def_data
    Next

    
    Erase CATEGORY_LIST
    
    If required_message <> "" Then
        required_message = "次の必須項目について入力されていない定義が存在します。" & required_message
        MsgBox required_message, vbOKOnly + vbExclamation, "入力エラー"
    End If
    
    
    Sheets("表紙").Select
    Cells(1, 1).Select
    
    Application.ScreenUpdating = True


End Sub


'# 過去履歴から調べたい配列を取得｡
'# 各シートの配列をfor each in で繰り返して、調べたい値を検索する。
'# 検索する際には filter(配列、要素、false) で存在しないものを探す。
'# 存在しないものはエラーを出力｡
'#
'# work は前回､今回を作成｡
'# 定義作成時に､今回を配列に格納し､前回にコピー｡
'#
'# 基本的に全て同じ処理｡
'# タグに複数定義する箇所と依存する定義はシート名をベースに追加処理｡
'#
'# hst , tgrp, job, snd, rcv, fmt, mfmt
'#
'# 削除も通知｡


