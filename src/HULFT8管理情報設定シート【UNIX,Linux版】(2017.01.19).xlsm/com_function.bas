Attribute VB_Name = "com_function"
'# __author__  = "Hiroshi Ohta"
'# __version__ = "0.01"
'# __date__    = "01 Nov 2019"

Function LengthCheck(ByVal CellsVal As String, ByVal char_limit As Integer) As String
        LengthCheck = ""
        If Len(CellsVal) > char_limit Then
            LengthCheck = CStr(char_limit) & " バイト以内で入力してください。"
        End If
End Function

Function RangeCheck(ByVal CellsVal As String, ByVal char_min_limit As Integer, ByVal char_max_limit As Integer) As String
        RangeCheck = ""
        If Len(CellsVal) < char_min_limit Or char_max_limit < Len(CellsVal) Then
            RangeCheck = CStr(char_min_limit) & " ～ " & CStr(char_max_limit) & " バイト以内で入力してください。"
        End If
End Function

Function StringCheck(ByVal CellsVal As String, ByVal check_pattern As String) As String

        char_limit = 1
        string_pattern = "[" & Replace(check_pattern, " ", ",") & "]"
        StringCheck = "'" & check_pattern & "' から1文字を入力してください。"

        If Len(CellsVal) <= char_limit Then
            If CellsVal Like string_pattern Or CellsVal = "" Then
                StringCheck = ""
            End If
        End If
End Function

Function StringRangeCheck(ByVal CellsVal As String, ByVal char_min_limit As Integer, ByVal char_max_limit As Long, ByVal prmit_zero As Integer) As String

        StringRangeCheck = ""
        prmit_zero_message = ""
        If CellsVal <> "" Then

            If IsNumeric(CellsVal) = False Then
                StringRangeCheck = "数値で入力してください。"
                Exit Function
            End If
            If prmit_zero = 0 And CellsVal = 0 Then
                Exit Function
            End If
            If CellsVal < char_min_limit Or char_max_limit < CellsVal Then
                If prmit_zero = 0 Then
                    prmit_zero_message = "0 または "
                End If

                StringRangeCheck = prmit_zero_message + CStr(char_min_limit) & " ～ " & CStr(char_max_limit) & " の範囲で入力してください。"
            End If
        End If
End Function

Function GetDefData(ByVal def_category As Variant) As Variant
    '# 定数定義
    '# -----------------------------------------------------------------------------------------------------------------
    Dim EXLS_MAX_COLUM As Long
    Dim EXLS_MAX_ROW As Long
    Dim CHECK_COLUMN As Integer
    
    '# 変数定義
    '# -----------------------------------------------------------------------------------------------------------------
    Dim column_cnt As Long
    Dim row_cnt As Long
    Dim data_column As Long
    Dim data_row As Long
    Dim check_row As Integer
    
    Dim add_column As Integer
    Dim str_data_row As Integer
    '# 定数設定
    '# -----------------------------------------------------------------------------------------------------------------
    EXLS_MAX_COLUM = 16384
    EXLS_MAX_ROW = 1048576
    
    
    '# パラメータ数取得
    check_row = 6
    If def_category = "hst" Then
        check_row = check_row + 1
    End If

    For column_cnt = 2 To EXLS_MAX_COLUM Step 1
        If Cells(check_row, column_cnt).Value <> "" Then
            data_column = column_cnt
        Else
            Exit For
        End If
    Next
        
    '# 登録レコード数取得
    CHECK_COLUMN = 2
    add_column = 0
    If def_category = "tgrp" Or def_category = "fmt" Then
        add_column = 1
    ElseIf def_category = "mfmt" Then
        add_column = 4
    End If
        
    For row_cnt = 12 To EXLS_MAX_COLUM Step 1
        If Cells(row_cnt, CHECK_COLUMN).Value <> "" Then
            data_row = row_cnt
        Else
            If Cells(row_cnt, CHECK_COLUMN + add_column).Value <> "" Then
                data_row = row_cnt
            Else
                Exit For
            End If
        End If
    Next
        
    '# 定義情報の取得
    str_data_row = check_row + 8
    If def_category = "tgrp" Then
        str_data_row = str_data_row + 1
    ElseIf InStr(def_category, "fmt") <> 0 Then
        str_data_row = str_data_row + 2
    End If
    
    Range(Cells(str_data_row, 2), Cells(data_row, data_column)).Sort key1:=Range(Cells(str_data_row, 2), Cells(data_row, 2)), _
                                                                order1:=xlAscending, _
                                                                Header:=xlNo, _
                                                                MatchCase:=True, _
                                                                SortMethod:=xlStroke
    GetDefData = Range(Cells(check_row, 2), Cells(data_row, data_column))


End Function


Function CheckDepsKeyDefined(ByVal def_data As Variant, ByVal check_deps_column As Integer, check_index_column As Integer, check_start_index_row As Integer) As String
    '# 定数定義
    '# -----------------------------------------------------------------------------------------------------------------
    Dim CHECK_DEPS_MAX As Integer
    
    '# 変数定義
    '# -----------------------------------------------------------------------------------------------------------------
    Dim check_deps_cnt As Integer
    Dim check_deps_data As Variant
    Dim check_deps_result As Variant
    Dim check_index_row As Integer

    
    CHECK_DEPS_MAX = 100
    ReDim check_deps_data(CHECK_DEPS_MAX)
    
    '# チェック対象のIDを配列に格納
    For check_deps_cnt = 0 To CHECK_DEPS_MAX
        check_deps_data(check_deps_cnt) = Cells(check_deps_cnt + 3, check_deps_column)
    Next
    
    CheckDepsKeyDefined = ""
    
    For check_index_row = check_start_index_row To UBound(def_data, 1) Step 1
        
        If def_data(check_index_row, check_index_column) <> "" Then
            '# check_deps_data 配列に対する検索結果を配列に格納
            check_deps_result = Filter(check_deps_data, def_data(check_index_row, check_index_column), True)
        
            '# 検索結果が存在しない場合は、チェックメッセージにチェック対象のIDを追加
            If UBound(check_deps_result) = -1 Then
                If InStr(CheckDepsKeyDefined, def_data(check_index_row, check_index_column)) = 0 Then
                    CheckDepsKeyDefined = CheckDepsKeyDefined & vbCrLf & "  - " & def_data(check_index_row, check_index_column)
                End If
            End If
            Erase check_deps_result
        End If
        
    Next
    
    If CheckDepsKeyDefined <> "" Then
        CheckDepsKeyDefined = "次の" & def_data(1, check_index_column) & "は、『" & check_deps_data(0) & "』に定義されてません。" & CheckDepsKeyDefined
    End If
    
    Erase check_deps_data

End Function

Sub HideWork()
    hist.Visible = xlSheetVeryHidden
End Sub

Sub ShowWork()
    hist.Visible = xlSheetVisible
    Sheets("登録履歴").Select
End Sub

