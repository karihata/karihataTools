Attribute VB_Name = "basListToSheet"
Option Explicit

Sub Show_frmListToSheet(ByVal control As IRibbonControl)
    frmListToSheet.Show
End Sub


Sub ListToSheet(Sheet_Name As String)
    Dim c As Range
    Dim WorkSheets_Dic As New Scripting.Dictionary
    Dim Selection_Dic As New Scripting.Dictionary
    Dim ws As Worksheet
    Dim Skip_Count As Long
    Dim Msg_Str As String
    Dim Sheet_Name_Error As Boolean
    Dim Named_Sheet_Name As String
    
    Skip_Count = 0
    
    '�V�[�g�쐬���Ƀ��[�N�V�[�g�Ƃ̃}�b�`���O���s�����߂Ɋ������[�N�V�[�g���m��
    For Each ws In ActiveWorkbook.Worksheets
        WorkSheets_Dic.Add ws.Name, 0
    Next ws
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '�I��͈͂ŋ󔒂̃Z���A�������͑I��͈͂ŏd�����Ă���ꍇ�͉������Ȃ�
    For Each c In Selection
        
        '�V�[�g�l�[���𕶎���œ��ꂷ�邽�߂ɕ�����^�̕ϐ��Ɋi�[
        Named_Sheet_Name = c.Value
        
        If Named_Sheet_Name <> "" And Not (Selection_Dic.Exists(Named_Sheet_Name)) Then
            
            '�I��͈͂ŏd�����Ȃ����m�F���邽�߂ɑI��͈͂̒l���m��
            Selection_Dic.Add Named_Sheet_Name, Named_Sheet_Name
            
            '�������[�N�V�[�g�ƃ}�b�`���O
            If WorkSheets_Dic.Exists(Named_Sheet_Name) Then
                '����������ɃX�L�b�v����\�����邽�߂ɃJ�E���g����
                Skip_Count = Skip_Count + 1
            Else
                '�A���}�b�`�̏ꍇ
                '�w�肵�����[�N�V�[�g�𖖔��ɃR�s�[���ăV�[�g����ύX����
                Worksheets(Sheet_Name).Copy after:=Worksheets(Worksheets.Count)
                
                On Error Resume Next
                
                Worksheets(Worksheets.Count).Name = Named_Sheet_Name
                
                '�V�[�g���̕ύX�Ɏ��s�����ꍇ�A�S�~�V�[�g���ł��Ă��܂����߁A�S�~�V�[�g���폜����
                If Worksheets(Worksheets.Count).Name <> Named_Sheet_Name Then
                    Worksheets(Worksheets.Count).Delete
                    Sheet_Name_Error = True
                End If
                
            End If
            
        End If
    Next c
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    '�I�u�W�F�N�g�����
    If Not (WorkSheets_Dic Is Nothing) Then Set WorkSheets_Dic = Nothing
    If Not (Selection_Dic Is Nothing) Then Set Selection_Dic = Nothing
    If Not (ws Is Nothing) Then Set ws = Nothing
    
    
    Msg_Str = "�������������܂����B"
    
    '�č쐬�t���O��FALSE�ŃX�L�b�v�񐔂�0��葽���ꍇ�̓X�L�b�v��\������
    If Skip_Count > 0 Then
         Msg_Str = Msg_Str & vbCrLf & "�G���[�P�F���ɓ����V�[�g��������܂����̂ŁA���̃V�[�g�쐬�̓X�L�b�v���܂����B"
    End If
    
    '�V�[�g���ύX�Ɏ��s�������̂����鋹�����b�Z�[�W�ɒǉ�����
    If Sheet_Name_Error Then
         Msg_Str = Msg_Str & vbCrLf & "�G���[�Q�F�V�[�g���Ƃ��Đݒ�ł��Ȃ����̂�����܂����̂ŁA���̃V�[�g�쐬�̓X�L�b�v���܂����B"
    End If
    
    '���b�Z�[�W��\��
    MsgBox Msg_Str
        
    Worksheets(Sheet_Name).Activate
    
End Sub
