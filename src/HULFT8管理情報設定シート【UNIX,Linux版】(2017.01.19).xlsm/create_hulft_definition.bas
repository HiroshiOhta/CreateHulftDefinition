Attribute VB_Name = "create_hulft_definition"
'# __author__  = "Hiroshi Ohta"
'# __version__ = "0.01"
'# __date__    = "01 Nov 2019"

Sub CreateHulftDefinition()

    Application.ScreenUpdating = False
    
    '# �萔��`
    '# -----------------------------------------------------------------------------------------------------------------
    Dim CATEGORY_LIST As Variant
    Dim HIST_MAX_CNT As Integer

    
    '# �萔�ݒ�
    '# -----------------------------------------------------------------------------------------------------------------
    CATEGORY_LIST = Array("hst", "tgrp", "job", "fmt", "mfmt", "snd", "rcv", "trg")
    HIST_MAX_CNT = 100


    '# �ϐ���`
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

    '# �ϐ��ݒ�
    '# -----------------------------------------------------------------------------------------------------------------
    out_def_dir = ActiveWorkbook.Path & "\def"

    '# -----------------------------------------------------------------------------------------------------------------
    '# �O����
    '# -----------------------------------------------------------------------------------------------------------------
    
    '# ���̓`�F�b�N
    '# -----------------------------------------------------------------------------------------------------------------
    Application.ScreenUpdating = True
    CheckRequired
    Application.ScreenUpdating = False
    
    hist.Visible = xlSheetVisible
    If hist.Cells(1, 24) = "NG" Then
        warning_message = "�G���[���������܂����B" & vbCrLf & vbCrLf & "��`�쐬�����͎��{�������܂���B" & vbCrLf & "���̓G���[���C�����čĎ��s���Ă��������B"
        MsgBox warning_message, vbOKOnly + vbExclamation, "���̓`�F�b�N�G���["
        GoTo Warning_Exit
    End If
    
    '# �O��o�^�f�[�^�̐ݒ�
    '# -----------------------------------------------------------------------------------------------------------------
    this_tm_column = UBound(CATEGORY_LIST) + 1
    
    Sheets("�o�^����").Select
    last_regd_data = hist.Range(Cells(3, this_tm_column + 1), Cells(HIST_MAX_CNT, this_tm_column * 2))
    hist.Range(Cells(4, this_tm_column + 1), Cells(HIST_MAX_CNT, this_tm_column * 2)).ClearContents
    hist.Range(Cells(3, 1), Cells(HIST_MAX_CNT, this_tm_column)) = last_regd_data
    
    '# -----------------------------------------------------------------------------------------------------------------
    '# ��`�쐬����
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
        
        Application.StatusBar = sheets_name & " �̒�`�t�@�C�����쐬���Ă��܂��B"
        
        '# ��`���̎擾
        '# -----------------------------------------------------------------------------------------------------------------
        def_data = GetDefData(def_category)
        
        out_def_file = out_def_dir & "\" & def_category & "_r.txt"
        
        '# ADODB.Stream�I�u�W�F�N�g�𐶐�
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
                .LineSeparator = 10 '# LF�ŏo��
                .Open
            
                '# �o�̓��R�[�h�̎擾
                insert_row = 4
                For index_row = start_index_row To UBound(def_data, 1) Step 1
                
                    '# �o�^�f�[�^��o�^�����V�[�g�֓���
                    insert_column = 8 + def_category_cnt
                    hist.Cells(insert_row, insert_column) = def_data(index_row, 1)
                    insert_row = insert_row + 1
                    
                    For index_column = LBound(def_data, 2) To UBound(def_data, 2) Step 1
                        
                        
                        If InStr(def_data(2, index_column), "�`") = 0 Then
                        '# KEY=VALUE �`���̒�`���`����
                        '# -----------------------------------------------------------------------------------------------------------------
                        
                            '# �o�̓��R�[�h�� KEY = VALUE �`���ŕϐ��ɕۑ�
                            If def_category = "rcv" And def_data(2, index_column) = "GENMNGNO" Then
                                def_line = def_data(2, index_column) & "=" & Format(def_data(index_row, index_column), "0000")
                            Else
                                def_line = def_data(2, index_column) & "=" & def_data(index_row, index_column)
                            End If

                            '# �o�̓��R�[�h�̐擪�f�[�^�̓w�b�_�[���o�͂���B
                            If index_column = LBound(def_data, 2) Then
                                def_line = "#" & vbLf & "# " & Replace(def_line, def_data(2, index_column), "ID") & vbLf & "#" & vbLf & vbLf & def_line
                            '# �Ō�̃��R�[�h�� END �����̍s�ɒǉ����ďo�͂���B
                            ElseIf index_column = UBound(def_data, 2) Then
                                def_line = def_line & vbLf & "END"
                            End If
                        
                        Else
                        '# �����s�A������Ɍׂ��`���`����
                        '# -----------------------------------------------------------------------------------------------------------------
                            
                            '# KEY�̊J�n�ƏI����z��Ɋi�[
                            keys = Split(def_data(2, index_column), "�`")
                            def_line = keys(0) & vbLf
                            
                            If InStr(def_category, "fmt") = 0 Then
                                def_line = def_line & " " & def_data(index_row, index_column)
                            Else
                                '# �t�H�[�}�b�g���ƃ}���`�t�H�[�}�b�g���̒�`
                                end_fmt_column = 6
                                
                                str_fmt_colum = 4
                                If def_category = "fmt" Then
                                    str_fmt_colum = 2
                                End If
                                
                                For next_column = str_fmt_colum To end_fmt_column Step 1
                                    '#
                                    '# Todo: �t�H�[�}�b�g���ƃ}���`�t�H�[�}�b�g���̏�����ǉ�
                                    '#
                                Next
                            End If
                            
                            def_line = def_line & vbLf
                            add_row = 0
                            
                            '# �����s�ɒl����`���邱�Ƃ��ł���
                            '# ���̂��߁A��`�̃L�[�ƂȂ�ID�����̍s�ɒ�`����Ă��邩���m�F���A��`����Ă��Ȃ��ꍇ�͓���ID�̃f�[�^�Ƃ��ď���
                            If index_row < UBound(def_data, 1) Then
                                For check_index_row = 1 To UBound(def_data) - start_index_row Step 1
                                    If def_data(index_row + check_index_row, LBound(def_data, 2)) = "" Then
                                    
                                    
                                        If InStr(def_category, "fmt") = 0 Then
                                            def_line = def_line & " " & def_data(index_row + check_index_row, index_column)
                                        Else
                                            '# �t�H�[�}�b�g���ƃ}���`�t�H�[�}�b�g���̒�`
                                            end_fmt_column = 6
                                            
                                            str_fmt_colum = 4
                                            If def_category = "fmt" Then
                                                str_fmt_colum = 2
                                            End If
                                            
                                            For next_column = str_fmt_colum To end_fmt_column Step 1
                                                '#
                                                '# Todo: �t�H�[�}�b�g���ƃ}���`�t�H�[�}�b�g���̏�����ǉ�
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
                    
                        '# �X�g���[���֏�����
                        .WriteText def_line, 1
                    Next
                
                    def_line = ""
                    .WriteText def_line, 1
                    
                    '# �����s���������ꍇ�A���������s�����C���f�b�N�X�������߂�B
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
    '# �O�񍷕��`�F�b�N(�폜 ����� /etc/hosts �ǉ��̊m�F)
    '# -----------------------------------------------------------------------------------------------------------------

    Application.StatusBar = "�폜���ꂽ ID �ƐV�K�o�^���ꂽ�z�X�g���`�F�b�N���Ă��܂��B"

    Sheets("�o�^����").Select
    this_regd_data = hist.Range(Cells(4, this_tm_column + 1), Cells(HIST_MAX_CNT, this_tm_column * 2))


    '# �V�K�ǉ����ꂽ�z�X�g�擾
    '# -----------------------------------------------------------------------------------------------------------------
    
    '# �O��o�^���ꂽ�z�X�g���X�g���쐬
    For index_row = LBound(last_regd_data, 1) To UBound(last_regd_data, 1) Step 1
        If last_regd_data(index_row, 1) <> "" Then
            ReDim Preserve last_regd_host(index_row - 1)
            last_regd_host(index_row - 1) = last_regd_data(index_row, 1)
        Else
            index_row = UBound(last_regd_data, 1)
        End If
    Next
    
    '# ����o�^�����z�X�g���O��o�^���X�g�ɑ��݂��邩���肵�A�ǉ��z�X�g�𒊏o
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
        add_host_message = "���̃z�X�g���ǉ�����Ă܂��B" & add_host_message & vbCrLf & vbCrLf & "/etc/hosts �̏C�������{���Ă��������B"
        MsgBox add_host_message, vbOKOnly + vbInformation
    End If
    
    Erase last_regd_host
    Erase check_result
    
    
    '# �폜���ꂽID���X�g�擾
    '# -----------------------------------------------------------------------------------------------------------------
    delete_message = ""
    For index_column = LBound(this_regd_data, 2) To UBound(this_regd_data, 2) Step 1
    
        '# �J�e�S���[���Ƀ`�F�b�N����o�^�������X�g���쐬
        For index_row = LBound(this_regd_data, 1) To UBound(this_regd_data, 1) Step 1
            If this_regd_data(index_row, index_column) <> "" Then
                ReDim Preserve this_regd_id(index_row - 1)
                this_regd_id(index_row - 1) = this_regd_data(index_row, index_column)
            Else
                index_row = UBound(this_regd_data, 1)
            End If
        Next

        '# �O��o�^����ID������o�^�������X�g�ɑ��݂��邩���肵�A�폜ID�𒊏o
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
        delete_message = "�O��o�^�����炢������ID���폜����Ă܂��B" & delete_message & vbCrLf & vbCrLf & "utlirm �R�}���h�ō폜�����{���Ă��������B"
        MsgBox delete_message, vbOKOnly + vbInformation
    End If


    Erase last_regd_data
    Erase this_regd_data
    
    Set adStrm = Nothing
    Set sht = Nothing

    Application.StatusBar = False

Warning_Exit:

    hist.Visible = xlSheetVeryHidden
    
    Sheets("�ڎ�").Select
    Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
End Sub
