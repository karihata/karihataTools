Attribute VB_Name = "basMicroStationCreateData"
Option Explicit

'�I��͈͓��́u�_����X��Y�v���}�C�N���X�e�[�V������XYZ�C���|�[�g�e�L�X�g���o�͂���
'�K���_���AX���W�AY���W�̗�̏��ԂɂȂ��Ă��邱��
Sub CreateXYText(ByVal control As IRibbonControl)
    If MsgBox("XY�e�L�X�g�t�@�C�����o�͂��܂��B" & vbCrLf & "�I��͈͂ɂ́u�_���v�uX���W�v�uY���W�v�̏��ԂɂȂ��Ă��܂����H" & vbCrLf & "���擪�s�͌��o���ł͂Ȃ��f�[�^������s�ɂ��ĉ������B" _
    , vbYesNo + vbQuestion, "�m�F") = vbYes Then
        
        '�I��͈͂�3��I�����Ă��Ȃ��ꍇ�̓G���[���b�Z�[�W��\�����ď��������Ȃ�
        If Selection.Columns.Count = 3 Then
            Dim Name_txt_File As Variant        '�o�͐�̓_��txt�t�@�C���p�X
            Dim Point_txt_File As Variant       '�o�͐�̓_txt�t�@�C���p�X
            Dim i As Long
            
            '������
            Name_txt_File = False
            Point_txt_File = False
            
            '���炩�̃G���[�����������ꍇ��ErrorTrap�̏��������s
            On Error GoTo ErrorTrap
            
            '�ۑ��t�@�C���������ꂼ�ꌈ�߂Ă��炤
            Name_txt_File = Application.GetSaveAsFilename(InitialFileName:=ActiveSheet.Name & "_PointName.txt", FileFilter:="txt�t�@�C��(*.txt),*.txt", Title:="�_��txt�t�@�C���̕ۑ���")
            Point_txt_File = Application.GetSaveAsFilename(InitialFileName:=ActiveSheet.Name & "_Point.txt", FileFilter:="txt�t�@�C��(*.txt),*.txt", Title:="�_txt�t�@�C���̕ۑ���")
            
            
            '�ۑ��t�@�C�����̂ǂ��炩��False�̏ꍇ�͏������s��Ȃ�
            If Name_txt_File <> False And Point_txt_File <> False Then
                
                '�_��txt�t�@�C�����I�[�v��
                Open Name_txt_File For Output As #1
                    
                    '�I�������s�̕������e�L�X�g�t�@�C���ɏ�������
                    For i = 0 To Selection.Rows.Count - 1
                        '�_������������
                        '�������ލۂɓ_���ɑS�p���p�X�y�[�X���܂܂�Ă���ꍇ�̓A���_�[�X�R�A�ɒu������
                        Print #1, Replace(Replace(Selection(1).Offset(i, 0).Value, " ", "_"), "�@", "_") & " " & Selection(1).Offset(i, 1).Value & " " & Selection(1).Offset(i, 2).Value & " " & "0" & vbCrLf;
                    Next i
                    
                Close #1
                
                Open Point_txt_File For Output As #1
                    
                    '�I�������s�̕������e�L�X�g�t�@�C���ɏ�������
                    For i = 0 To Selection.Rows.Count - 1
                        '�_������������
                        Print #1, Selection(1).Offset(i, 1).Value & " " & Selection(1).Offset(i, 2).Value & " " & "0" & vbCrLf;
                    Next i
                    
                Close #1
                
                MsgBox "�e�L�X�g�t�@�C�����o�͂��܂����B"
            End If

        Else
            MsgBox "�u�_���v�uX���W�v�uY���W�v��3���I�����ĉ������B"
        
        End If
        
        

    End If
    
Exit Sub
ErrorTrap:
    MsgBox "�G���[���������܂����B" & vbCrLf & "�G���[�ԍ��F" & Err.Number & vbCrLf & "�G���[���e" & vbCrLf & Err.Description
    Close #1
End Sub
