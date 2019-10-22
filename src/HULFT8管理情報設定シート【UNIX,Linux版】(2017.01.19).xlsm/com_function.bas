Attribute VB_Name = "com_function"
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

                StringRangeCheck = prmit_zero_message + CStr(char_min_limit) & " ～ " & CStr(char_max_limit) & " バイト以内で入力してください。"
            End If
        End If
End Function
