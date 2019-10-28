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
    
    
    '# 変数定義
    '# -----------------------------------------------------------------------------------------------------------------
    Dim column_cnt As Long
    Dim row_cnt As Long
    Dim data_column As Long
    Dim data_row As Long
    Dim check_row As Integer
    Dim def_data As Variant
    
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
    check_column = 2
    If def_category = "tgrp" Or def_category = "fmt" Then
        check_column = check_column + 1
    ElseIf def_category = "mfmt" Then
        check_column = check_column + 4
    End If
        
    For row_cnt = 12 To EXLS_MAX_COLUM Step 1
        If Cells(row_cnt, check_column).Value <> "" Then
            data_row = row_cnt
        Else
            Exit For
        End If
    Next
        
    '# 定義情報の取得
    GetDefData = Range(Cells(check_row, 2), Cells(data_row, data_column))


End Function
