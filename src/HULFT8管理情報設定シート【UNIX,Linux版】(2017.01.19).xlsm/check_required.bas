Attribute VB_Name = "check_required"
'# __author__  = "Hiroshi Ohta"
'# __version__ = "0.01"
'# __date__    = "01 Nov 2019"

Sub CheckRequired()

    Application.ScreenUpdating = False
    
    '# �萔��`
    '# -----------------------------------------------------------------------------------------------------------------
    Dim CATEGORY_LIST As Variant
        
    Dim EXLS_MAX_COLUM As Long
    Dim EXLS_MAX_ROW As Long
    Dim CHECK_DEPS_MAX As Integer
    
    
    '# �ϐ���`
    '# -----------------------------------------------------------------------------------------------------------------
    Dim sht As Worksheet
    
    Dim check_status As String
    Dim def_category As Variant
    Dim sheets_name As String
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
    Dim check_start_index_row As Integer
    Dim check_index_column As Integer
    Dim check_deps_message As String
    
    
    
    '# �萔�ݒ�
    '# -----------------------------------------------------------------------------------------------------------------
    CATEGORY_LIST = Array("hst", "tgrp", "job", "fmt", "mfmt", "snd", "rcv", "trg")
    CHECK_DEPS_MAX = 100
    
    
    '# �ϐ��ݒ�
    '# -----------------------------------------------------------------------------------------------------------------
    check_status = "OK"
    required_message = ""
    def_category_cnt = 0
    hist_row = 4
    
    '# �O����
    '# -----------------------------------------------------------------------------------------------------------------
    hist.Visible = xlSheetVisible
    hist.Select
    Range(Cells(hist_row, 17), Cells(CHECK_DEPS_MAX, 17 + 8)).ClearContents
    
    '# ��`�`�F�b�N
    '# -----------------------------------------------------------------------------------------------------------------
    For Each def_category In CATEGORY_LIST
        def_category_cnt = def_category_cnt + 1
        '# �Y���V�[�g�̎擾
        For Each sht In Worksheets
            If def_category = sht.CodeName Then
                sheets_name = sht.Name
                Sheets(sheets_name).Select
                Exit For
            End If
        Next
        
        Application.StatusBar = sheets_name & " �̓��͂��ꂽ�l���`�F�b�N���Ă��܂��B"

        
        '# ��`���̎擾
        '# -----------------------------------------------------------------------------------------------------------------
        def_data = GetDefData(def_category)
        
        
        '# �K�{���ڂ̊m�F
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
            If def_data(check_required, check_index_column) = "��" Then
                error_cnt = 0
                For check_index_row = check_start_index_row To UBound(def_data, 1) Step 1
                    If def_data(check_index_row, check_index_column) = "" Then
                        
                        '# ���ID�ŕ����̕����s�̒�`���쐬����ꍇ
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
            required_message = required_message & vbCrLf & vbCrLf & "�V�[�g���F" & sheets_name & required_key
        End If
        
        
        '# �ˑ����ڂ̊m�F
        '# -----------------------------------------------------------------------------------------------------------------
        '# tgrp �� hst, mfmt �� fmt, snd �� job tgrp fmt mfmt, rcv �� job tgrp, trg �� job
        hist.Select
        check_deps_message = ""
        If def_category <> "hst" And def_category <> "job" And def_category <> "fmt" Then
        
            '# �]���O���[�v�̊m�F
            '# -----------------------------------------------------------------------------------------------------------------
            If def_category = "tgrp" Then

                check_deps_column = 17
                check_index_column = 2  '# �z�X�g�� �̔z�񎟐�
                check_deps_message = CheckDepsKeyDefined(def_data, check_deps_column, check_index_column, check_start_index_row)
                    
                If check_deps_message <> "" Then
                    check_deps_message = "�V�[�g���F" & sheets_name & vbCrLf & vbCrLf & check_deps_message
                    MsgBox check_deps_message, vbOKOnly + vbExclamation, "���̓G���["
                    check_status = "NG"
                End If
            
            
            '# �z�M�Ǘ����̊m�F
            '# -----------------------------------------------------------------------------------------------------------------
            ElseIf def_category = "snd" Then
            
                '# �]���O���[�v�̊m�F
                check_deps_column = 18
                check_index_column = 19
                check_deps_message = CheckDepsKeyDefined(def_data, check_deps_column, check_index_column, check_start_index_row)
                
                If check_deps_message <> "" Then
                    check_deps_message = "�V�[�g���F" & sheets_name & vbCrLf & vbCrLf & check_deps_message
                    MsgBox check_deps_message, vbOKOnly + vbExclamation, "���̓G���["
                    check_status = "NG"
                End If
                
                '# �W���u�N���̊m�F
                check_deps_column = 19
                For check_index_column = 15 To 17 Step 1
                    check_deps_message = CheckDepsKeyDefined(def_data, check_deps_column, check_index_column, check_start_index_row)
                    
                    If check_deps_message <> "" Then
                        check_deps_message = "�V�[�g���F" & sheets_name & vbCrLf & vbCrLf & check_deps_message
                        MsgBox check_deps_message, vbOKOnly + vbExclamation, "���̓G���["
                        check_status = "NG"
                    End If
                Next
                
                '# �t�H�[�}�b�g�̊m�F(TRANSTYPE��F�̎�)
                '# check_deps_column = 20
                
                '# �}���`�t�H�[�}�b�g�̊m�F(TRANSTYPE��M�̎�)
                '# check_deps_column = 21
                
            
            '# �W�M�Ǘ����̊m�F
            '# -----------------------------------------------------------------------------------------------------------------
            ElseIf def_category = "rcv" Then
            
                '# �]���O���[�v�̊m�F
                check_deps_column = 18
                check_index_column = 14
                check_deps_message = CheckDepsKeyDefined(def_data, check_deps_column, check_index_column, check_start_index_row)
                    
                If check_deps_message <> "" Then
                    check_deps_message = "�V�[�g���F" & sheets_name & vbCrLf & vbCrLf & check_deps_message
                    MsgBox check_deps_message, vbOKOnly + vbExclamation, "���̓G���["
                    check_status = "NG"
                End If
                
                '# �W���u�N���̊m�F
                check_deps_column = 19
                For check_index_column = 12 To 13 Step 1
                    check_deps_message = CheckDepsKeyDefined(def_data, check_deps_column, check_index_column, check_start_index_row)
                    
                    If check_deps_message <> "" Then
                        check_deps_message = "�V�[�g���F" & sheets_name & vbCrLf & vbCrLf & check_deps_message
                        MsgBox check_deps_message, vbOKOnly + vbExclamation, "���̓G���["
                        check_status = "NG"
                    End If
                Next
                
                '# �t�H�[�}�b�g�̊m�F(TRANSTYPE��F�̎�)
                '# check_deps_column = 20
                
                '# �}���`�t�H�[�}�b�g�̊m�F(TRANSTYPE��M�̎�)
                '# check_deps_column = 21

            End If
        End If
        
        
        '# ����o�^�V�[�g�̃`�F�b�N�փf�[�^����
        '# -----------------------------------------------------------------------------------------------------------------
        insert_column = def_category_cnt + 16
        insert_row = hist_row
        For check_index_row = check_start_index_row To UBound(def_data, 1) Step 1
            If def_data(check_index_row, 1) <> "" Then
                Cells(insert_row, insert_column) = def_data(check_index_row, 1)
                insert_row = insert_row + 1
            End If
        Next
        
        
        '# �z��̊J��
        Erase def_data
        Application.StatusBar = False
    Next
    
    Erase CATEGORY_LIST
    
    If required_message <> "" Then
        required_message = "���̕K�{���ڂɂ��ē��͂���Ă��Ȃ���`�����݂��܂��B" & required_message
        MsgBox required_message, vbOKOnly + vbExclamation, "���̓G���["
        check_status = "NG"
    End If
    
    Set sht = Nothing
    Application.StatusBar = False
    
    hist.Cells(1, 24) = check_status
    hist.Visible = xlSheetVeryHidden
    
    Sheets("�ڎ�").Select
    Cells(1, 1).Select
    
    Application.ScreenUpdating = True

End Sub
