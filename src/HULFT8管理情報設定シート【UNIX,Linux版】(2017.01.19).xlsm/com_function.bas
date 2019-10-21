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
        StringCheck = check_pattern & " から1文字を入力してください。"

        If Len(CellsVal) <= char_limit Then
            If CellsVal Like string_pattern Or CellsVal = "" Then
                StringCheck = ""
            End If
        End If
End Function
