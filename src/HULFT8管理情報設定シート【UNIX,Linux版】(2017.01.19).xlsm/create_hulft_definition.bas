Attribute VB_Name = "create_hulft_definition"
'# __author__  = "Hiroshi Ohta"
'# __version__ = "0.01"
'# __date__    = "01 Nov 2019"

Sub CreateHulftDefinition()

    Application.ScreenUpdating = False
    
    '# 定数定義
    '# -----------------------------------------------------------------------------------------------------------------
    Dim CATEGORY_LIST As Variant
    Dim HIST_MAX_CNT As Integer

    
    '# 定数設定
    '# -----------------------------------------------------------------------------------------------------------------
    CATEGORY_LIST = Array("hst", "tgrp", "job", "fmt", "mfmt", "snd", "rcv", "trg")
    HIST_MAX_CNT = 100


    '# 変数定義
    '# -----------------------------------------------------------------------------------------------------------------
    Dim warning_message As String
    Dim def_category As Variant
    Dim def_category_cnt As Integer
    Dim sht As Worksheet
    Dim sheets_name As String
    
    Dim out_def_dir As String
    Dim out_def_file As String
    Dim start_index_row As String
    Dim index_row As Integer
    Dim index_column As Integer
    
    Dim def_line As String
    Dim keys As Variant
    Dim check_index_row As Integer
    Dim check_index_column As Integer

    Dim last_regd_data As Variant
    Dim this_regd_data As Variant
    
    Dim last_regd_host() As Variant
    Dim this_regd_id() As Variant
    Dim check_result As Variant

    '# 変数設定
    '# -----------------------------------------------------------------------------------------------------------------
    out_def_dir = ActiveWorkbook.Path & "\def"

    '# -----------------------------------------------------------------------------------------------------------------
    '# 前処理
    '# -----------------------------------------------------------------------------------------------------------------
    
    '# 入力チェック
    '# -----------------------------------------------------------------------------------------------------------------
    Application.ScreenUpdating = True
    CheckRequired
    Application.ScreenUpdating = False
    
    hist.Visible = xlSheetVisible
    If hist.Cells(1, 24) = "NG" Then
        warning_message = "エラーが発生しました。" & vbCrLf & vbCrLf & "定義作成処理は実施いたしません。" & vbCrLf & "入力エラーを修正して再実行してください。"
        MsgBox warning_message, vbOKOnly + vbExclamation, "入力チェックエラー"
        GoTo Warning_Exit
    End If
    
    '# 前回登録データの設定
    '# -----------------------------------------------------------------------------------------------------------------
    this_tm_column = UBound(CATEGORY_LIST) + 1
    
    Sheets("登録履歴").Select
    last_regd_data = hist.Range(Cells(3, this_tm_column + 1), Cells(HIST_MAX_CNT, this_tm_column * 2))
    hist.Range(Cells(4, this_tm_column + 1), Cells(HIST_MAX_CNT, this_tm_column * 2)).ClearContents
    hist.Range(Cells(3, 1), Cells(HIST_MAX_CNT, this_tm_column)) = last_regd_data
    
    '# -----------------------------------------------------------------------------------------------------------------
    '# 定義作成処理
    '# -----------------------------------------------------------------------------------------------------------------
    For Each def_category In CATEGORY_LIST
        def_category_cnt = def_category_cnt + 1
        '# 該当シートの取得
        For Each sht In Worksheets
            If def_category = sht.CodeName Then
                sheets_name = sht.Name
                Sheets(sheets_name).Select
                Exit For
            End If
        Next
        
        Application.StatusBar = sheets_name & " の定義ファイルを作成しています。"
        
        '# 定義情報の取得
        '# -----------------------------------------------------------------------------------------------------------------
        def_data = GetDefData(def_category)
        
        out_def_file = out_def_dir & "\" & def_category & "_r.txt"
        
        '# ADODB.Streamオブジェクトを生成
        Dim adStrm As Object
        Set adStrm = CreateObject("ADODB.Stream")
        
        start_index_row = 9
        If def_category = "tgrp" Then
            start_index_row = start_index_row + 1
        ElseIf def_category = "fmt" Or def_category = "mfmt" Then
            start_index_row = start_index_row + 2
        End If
        
        If start_index_row < UBound(def_data) Then
            With adStrm
                .Charset = "UTF-8"
                .LineSeparator = 10 '# LFで出力
                .Open
            
                '# 出力レコードの取得
                insert_row = 4
                For index_row = start_index_row To UBound(def_data, 1) Step 1
                
                    '# 登録データを登録履歴シートへ入力
                    insert_column = 8 + def_category_cnt
                    hist.Cells(insert_row, insert_column) = def_data(index_row, 1)
                    insert_row = insert_row + 1
                    
                    For index_column = LBound(def_data, 2) To UBound(def_data, 2) Step 1
                        
                        
                        If InStr(def_data(2, index_column), "～") = 0 Then
                        '# KEY=VALUE 形式の定義整形処理
                        '# -----------------------------------------------------------------------------------------------------------------
                        
                            '# 出力レコードを KEY = VALUE 形式で変数に保存
                            If def_category = "rcv" And def_data(2, index_column) = "GENMNGNO" Then
                                def_line = def_data(2, index_column) & "=" & Format(def_data(index_row, index_column), "0000")
                            Else
                                def_line = def_data(2, index_column) & "=" & def_data(index_row, index_column)
                            End If

                            '# 出力レコードの先頭データはヘッダーを出力する。
                            If index_column = LBound(def_data, 2) Then
                                def_line = "#" & vbLf & "# " & Replace(def_line, def_data(2, index_column), "ID") & vbLf & "#" & vbLf & vbLf & def_line
                            '# 最後のレコードは END を次の行に追加して出力する。
                            ElseIf index_column = UBound(def_data, 2) Then
                                def_line = def_line & vbLf & "END"
                            End If
                        
                        Else
                        '# 複数行、複数列に跨る定義整形処理
                        '# -----------------------------------------------------------------------------------------------------------------
                            
                            '# KEYの開始と終了を配列に格納
                            keys = Split(def_data(2, index_column), "～")
                            def_line = keys(0) & vbLf
                            
                            If InStr(def_category, "fmt") = 0 Then
                                def_line = def_line & " " & def_data(index_row, index_column)
                            Else
                                '# フォーマット情報とマルチフォーマット情報の定義
                                end_fmt_column = 6
                                
                                str_fmt_colum = 4
                                If def_category = "fmt" Then
                                    str_fmt_colum = 2
                                End If
                                
                                For next_column = str_fmt_colum To end_fmt_column Step 1
                                    '#
                                    '# Todo: フォーマット情報とマルチフォーマット情報の処理を追加
                                    '#
                                Next
                            End If
                            
                            def_line = def_line & vbLf
                            add_row = 0
                            
                            '# 複数行に値が定義することができる
                            '# そのため、定義のキーとなるIDが次の行に定義されているかを確認し、定義されていない場合は同一IDのデータとして処理
                            If index_row < UBound(def_data, 1) Then
                                For check_index_row = 1 To UBound(def_data) - start_index_row Step 1
                                    If def_data(index_row + check_index_row, LBound(def_data, 2)) = "" Then
                                    
                                    
                                        If InStr(def_category, "fmt") = 0 Then
                                            def_line = def_line & " " & def_data(index_row + check_index_row, index_column)
                                        Else
                                            '# フォーマット情報とマルチフォーマット情報の定義
                                            end_fmt_column = 6
                                            
                                            str_fmt_colum = 4
                                            If def_category = "fmt" Then
                                                str_fmt_colum = 2
                                            End If
                                            
                                            For next_column = str_fmt_colum To end_fmt_column Step 1
                                                '#
                                                '# Todo: フォーマット情報とマルチフォーマット情報の処理を追加
                                                '#
                                            Next
                                        End If
                                        
                                        def_line = def_line & vbLf
                                        add_row = check_index_row
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                            
                            def_line = def_line & keys(1)
                            Erase keys
                        End If
                    
                        '# ストリームへ書込み
                        .WriteText def_line, 1
                    Next
                
                    def_line = ""
                    .WriteText def_line, 1
                    
                    '# 複数行処理した場合、処理した行数分インデックスをすすめる。
                    index_row = index_row + add_row
                Next
            
                .SaveToFile out_def_file, 2
                .Close
            End With
        End If
        
        Erase def_data
        Application.StatusBar = False

    Next
    Erase CATEGORY_LIST
    
    '# -----------------------------------------------------------------------------------------------------------------
    '# 前回差分チェック(削除 および /etc/hosts 追加の確認)
    '# -----------------------------------------------------------------------------------------------------------------

    Application.StatusBar = "削除された ID と新規登録されたホストをチェックしています。"

    Sheets("登録履歴").Select
    this_regd_data = hist.Range(Cells(4, this_tm_column + 1), Cells(HIST_MAX_CNT, this_tm_column * 2))


    '# 新規追加されたホスト取得
    '# -----------------------------------------------------------------------------------------------------------------
    
    '# 前回登録されたホストリストを作成
    For index_row = LBound(last_regd_data, 1) To UBound(last_regd_data, 1) Step 1
        If last_regd_data(index_row, 1) <> "" Then
            ReDim Preserve last_regd_host(index_row - 1)
            last_regd_host(index_row - 1) = last_regd_data(index_row, 1)
        Else
            index_row = UBound(last_regd_data, 1)
        End If
    Next
    
    '# 今回登録したホストが前回登録リストに存在するか判定し、追加ホストを抽出
    add_host_message = ""
    For index_row = LBound(this_regd_data, 1) To UBound(this_regd_data, 1) Step 1
        If this_regd_data(index_row, 1) <> "" Then
            check_result = Filter(last_regd_host, this_regd_data(index_row, 1), True)
            If UBound(check_result) = -1 Then
                add_host_message = add_host_message & vbCrLf & "  - " & this_regd_data(index_row, 1)
            End If
        Else
            index_row = UBound(this_regd_data, 1)
        End If
    Next
    
    If add_host_message <> "" Then
        add_host_message = "次のホストが追加されてます。" & add_host_message & vbCrLf & vbCrLf & "/etc/hosts の修正も実施してください。"
        MsgBox add_host_message, vbOKOnly + vbInformation
    End If
    
    Erase last_regd_host
    Erase check_result
    
    
    '# 削除されたIDリスト取得
    '# -----------------------------------------------------------------------------------------------------------------
    delete_message = ""
    For index_column = LBound(this_regd_data, 2) To UBound(this_regd_data, 2) Step 1
    
        '# カテゴリー毎にチェック今回登録したリストを作成
        For index_row = LBound(this_regd_data, 1) To UBound(this_regd_data, 1) Step 1
            If this_regd_data(index_row, index_column) <> "" Then
                ReDim Preserve this_regd_id(index_row - 1)
                this_regd_id(index_row - 1) = this_regd_data(index_row, index_column)
            Else
                index_row = UBound(this_regd_data, 1)
            End If
        Next

        '# 前回登録したIDが今回登録したリストに存在するか判定し、削除IDを抽出
        del_id_message = ""
        For index_row = LBound(last_regd_data, 1) + 1 To UBound(last_regd_data, 1) Step 1
            If last_regd_data(index_row, index_column) <> "" Then
                check_result = Filter(this_regd_id, last_regd_data(index_row, index_column), True)
                If UBound(check_result) = -1 Then
                    del_id_message = del_id_message & vbCrLf & "  - " & last_regd_data(index_row, index_column)
                End If
            Else
                index_row = UBound(last_regd_data, 1)
            End If
        Next
        
        If del_id_message <> "" Then
            delete_message = delete_message & vbCrLf & vbCrLf & last_regd_data(1, index_column) & del_id_message
        End If
        
        Erase this_regd_id
        Erase check_result

    Next
    
    If delete_message <> "" Then
        delete_message = "前回登録分からいくつかのIDが削除されてます。" & delete_message & vbCrLf & vbCrLf & "utlirm コマンドで削除を実施してください。"
        MsgBox delete_message, vbOKOnly + vbInformation
    End If


    Erase last_regd_data
    Erase this_regd_data
    
    Set adStrm = Nothing
    Set sht = Nothing

    Application.StatusBar = False

Warning_Exit:

    hist.Visible = xlSheetVeryHidden
    
    Sheets("目次").Select
    Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
End Sub
