Attribute VB_Name = "basCountImageResolution"
Option Explicit

Sub Count_Image_Resolution_Powered_By_ImageMagick(ByVal control As IRibbonControl)
    Dim c As Range
    Dim i As Long
    Dim img As Object
    Dim objpb
    
    '�i���Ǘ��̂��߂ɁA0�ŏ�����
    i = 0
    
    If MsgBox("�I���Z���̉E��2�񕪂̃Z���ɏc����DPI���i�[���܂��B" & vbCrLf & "��낵���ł����B", vbYesNo) = vbYes Then
        '�I��͈͂�1��̏ꍇ�̂ݏ���������
        If Selection.Columns.Count = 1 Then
            
            ' �i���o�[�N���X�̒�`
            Set objpb = New ProgressBar
            
            ' �i���o�[�̃^�C�g����ݒ�
            objpb.SetTitle "�h�L�������g��DPI�m�F"

            '�摜�ȊO�̂��̂��o�Ă��Ă��G���[�Ŏ~�܂�Ȃ��悤�ɂ���
            On Error Resume Next
        
            'ImageMagick���g�����߂ɃI�u�W�F�N�g���쐬
            Set img = New ImageMagickObject.MagickImage
            
            For Each c In Selection
                '�Y���Z�����󔒂̏ꍇ�͉������Ȃ�
                If c.Value <> "" Then
                    c.Offset(0, 1).Value = Split(img.Identify("-format", "%y,", c.Value), ",")(0)
                    c.Offset(0, 2).Value = Split(img.Identify("-format", "%x,", c.Value), ",")(0)
                End If
                
                '�i���̃J�E���g�A�b�v
                i = i + 1
                
                '�v���O���X�o�[�X�V
                objpb.SetTitle "�h�L�������g��DPI�m�F�@" & i & "�F" & Selection.Count
                objpb.SetProgress i / Selection.Count
            Next
            
            Set img = Nothing
                    
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


