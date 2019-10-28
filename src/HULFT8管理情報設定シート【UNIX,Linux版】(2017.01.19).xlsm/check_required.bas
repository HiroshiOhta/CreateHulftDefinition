Attribute VB_Name = "check_required"
'# __author__  = "Hiroshi Ohta"
'# __version__ = "0.01"
'# __date__    = "01 Nov 2019"

Sub CheckRequired()

    Application.ScreenUpdating = False
    
    '# �萔��`
    '# -----------------------------------------------------------------------------------------------------------------
    Dim CATEGORY_LIST As Variant
    Dim SH As Worksheet
    
    
    Dim EXLS_MAX_COLUM As Long
    Dim EXLS_MAX_ROW As Long
    Dim CHECK_DEPS_MAX As Integer
    
    
    '# �ϐ���`
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
    
    
    '# �萔�ݒ�
    '# -----------------------------------------------------------------------------------------------------------------
    CATEGORY_LIST = Array("hst", "tgrp", "job", "fmt", "mfmt", "snd", "rcv", "trg")
    EXLS_MAX_COLUM = 16384
    EXLS_MAX_ROW = 1048576
    
    CHECK_DEPS_MAX = 100
    
    
    '# ��`�`�F�b�N
    '# -----------------------------------------------------------------------------------------------------------------
    required_message = ""
    def_category_cnt = 0
    
    '# �O����
    '# -----------------------------------------------------------------------------------------------------------------
    hist.Select
    Range(Cells(3, 17), Cells(CHECK_DEPS_MAX, 17 + 8)).ClearContents
    
    For Each def_category In CATEGORY_LIST
        def_category_cnt = def_category_cnt + 1
        '# �Y���V�[�g�̎擾
        For Each SH In Worksheets
            If def_category = SH.CodeName Then
                Sheets(SH.Name).Select
                Exit For
            End If
        Next
        
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
            required_message = required_message & vbCrLf & vbCrLf & "�V�[�g���F" & SH.Name & required_key
        End If
        
        
        '# �ˑ����ڂ̊m�F
        '# -----------------------------------------------------------------------------------------------------------------
        '# tgrp �� hst, mfmt �� fmt, snd �� job tgrp fmt mfmt, rcv �� job tgrp, trg �� job
        hist.Select
        If def_category <> "hst" And def_category <> "job" And def_category <> "fmt" Then
        
            '# �]���O���[�v�̊m�F
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
                    check_deps_message = "����" & def_data(1, check_index_column) & "�́A�ڍ׃z�X�g���ɒ�`����Ă܂���B" & check_deps_message
                    MsgBox check_deps_message, vbOKOnly + vbExclamation, "���̓G���["
                End If
                
                Erase check_deps_data
                Erase check_deps_result
                
            End If
        End If
        
        
        
        
        '# ����o�^�V�[�g�̃`�F�b�N�փf�[�^����
        '# -----------------------------------------------------------------------------------------------------------------
        insert_column = def_category_cnt + 16
        
        insert_row = 3
        For check_index_row = check_start_index_row To UBound(def_data, 1) Step 1
            If def_data(check_index_row, 1) <> "" Then
                Cells(insert_row, insert_column) = def_data(check_index_row, 1)
                insert_row = insert_row + 1
            End If
        Next
        
        
        '# �z��̊J��
        Erase def_data
    Next

    
    Erase CATEGORY_LIST
    
    If required_message <> "" Then
        required_message = "���̕K�{���ڂɂ��ē��͂���Ă��Ȃ���`�����݂��܂��B" & required_message
        MsgBox required_message, vbOKOnly + vbExclamation, "���̓G���["
    End If
    
    
    Sheets("�\��").Select
    Cells(1, 1).Select
    
    Application.ScreenUpdating = True


End Sub


'# �ߋ��������璲�ׂ����z����擾�
'# �e�V�[�g�̔z���for each in �ŌJ��Ԃ��āA���ׂ����l����������B
'# ��������ۂɂ� filter(�z��A�v�f�Afalse) �ő��݂��Ȃ����̂�T���B
'# ���݂��Ȃ����̂̓G���[���o�͡
'#
'# work �͑O�񤍡����쐬�
'# ��`�쐬���ɤ�����z��Ɋi�[����O��ɃR�s�[�
'#
'# ��{�I�ɑS�ē��������
'# �^�O�ɕ�����`����ӏ��ƈˑ������`�̓V�[�g�����x�[�X�ɒǉ������
'#
'# hst , tgrp, job, snd, rcv, fmt, mfmt
'#
'# �폜���ʒm�


