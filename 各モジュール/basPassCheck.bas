Attribute VB_Name = "basPassCheck"
Option Explicit

Sub Path_Check(ByVal control As IRibbonControl)
    Dim c As Range
    Dim i As Long
    Dim objpb
    
    '�i���Ǘ��̂��߂ɁA0�ŏ�����
    i = 0
    
    If MsgBox("�I���Z���̉E�ׂ̃Z���Ɍ��ʂ��i�[���܂��B" & vbCrLf & "��낵���ł����B", vbYesNo) = vbYes Then
        '�I��͈͂�1��̏ꍇ�̂ݏ���������
        If Selection.Columns.Count = 1 Then
            
            ' �i���o�[�N���X�̒�`
            Set objpb = New ProgressBar
            
            ' �i���o�[�̃^�C�g����ݒ�
            objpb.SetTitle "�p�X�m�F"
            
            '�I��͈͂�����
            For Each c In Selection
                
                On Error Resume Next
                
                '�I���Z�����u�����N�ȊO�̏ꍇ�͏����𑱂���
                If c.Value <> "" Then
                    c.Offset(0, 1).Value = FolderExists(c.Value)
                Else
                    '�I���Z�����u�����N�̏ꍇ�͉������Ȃ�
                    c.Offset(0, 1).Value = ""
                End If
                
                '�i���̃J�E���g�A�b�v
                i = i + 1
                
                '�v���O���X�o�[�X�V
                objpb.SetTitle "�p�X�m�F�@" & i & "�F" & Selection.Count
                objpb.SetProgress i / Selection.Count
            Next c
            
            '�v���O���X�o�[�̔j��
            Set objpb = Nothing
            
            MsgBox "�������������܂����B"
        Else
            '�I��͈͂�1��ł͂Ȃ������ꍇ�̓G���[���b�Z�[�W��\��
            MsgBox "1�񂾂��I�����Ă��������B"
        End If
    End If
    
    '�X�e�[�^�X�o�[�����Z�b�g
    Application.StatusBar = False
    
End Sub

Function FolderExists(Path_String As String) As Boolean
  'FileSystemObject�Ńt�@�C�����݃`�F�b�N������
  Dim fso As New Scripting.FileSystemObject
  
  '���炩�̃G���[�����������ꍇ��ErrorTrap�̏��������s
  On Error GoTo ErrorTrap
  
  If fso.FileExists(Path_String) Then
    FolderExists = True
  Else
    FolderExists = False
  End If
    
  '�I�u�W�F�N�g�̕Еt��
  Set fso = Nothing
  
Exit Function

ErrorTrap:
    If Not (fso Is Nothing) Then Set fso = Nothing
    FolderExists = False
End Function
