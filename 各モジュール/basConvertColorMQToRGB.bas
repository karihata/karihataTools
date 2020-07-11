Attribute VB_Name = "basConvertColorMQToRGB"
Option Explicit

Sub ConvertColorMQToRGB(ByVal control As IRibbonControl)
    Dim c As Range
    Dim i As Long
    Dim R As Long
    Dim G As Long
    Dim B As Long
    Dim inputMQcolor As String
    Dim objpb
    
    '�i���Ǘ��̂��߂ɁA0�ŏ�����
    i = 0
    
    If MsgBox("�I��͈͂��Z����MQ�J���[�l�œh��Ԃ��܂��B" & vbCrLf & "��낵���ł����B" & vbCrLf & vbCrLf & "�u-�v(�n�C�t��)�̃Z���ɂ͉������܂���B", vbYesNo) = vbYes Then
        ' �i���o�[�N���X�̒�`
        Set objpb = New ProgressBar
        
        ' �i���o�[�̃^�C�g����ݒ�
        objpb.SetTitle "MQ�J���[�ŃZ���̐F��ύX��"
        
        '�I��͈͂�����
        For Each c In Selection
            
            '�f�o�b�O�p�Ɋi�[
            inputMQcolor = c.Value
            
            On Error Resume Next
            
            '�I���Z���̒l���u�����N�ł͂Ȃ����ƁA"-"�ɂȂ��Ă��Ȃ����Ƃ��m�F
            '�Y�������ꍇ�͉������Ȃ�
            If Not (IsEmpty(c.Value)) And inputMQcolor <> "-" Then
                
                '�I���Z���̒l�̐����`�F�b�N
                '�����ł͂Ȃ��ꍇ�͉������Ȃ�
                If IsNumeric(inputMQcolor) And c.Value = Int(inputMQcolor) Then
                    '�Z���̒l����MQ�J���[�ԍ�����RGB�̔ԍ��ɂ��ꂼ��ϊ����Ċi�[
                    '256 * 256 = 65536
                    B = inputMQcolor \ 65536
                    
                    G = (inputMQcolor - (B * 65536)) \ 256
                    
                    R = inputMQcolor - (G * 256) - (B * 65536)
                                  
                    '�Z���̐F�ύX
                    c.Interior.Color = RGB(R, G, B)
                End If
            
            End If
            
            '�i���̃J�E���g�A�b�v
            i = i + 1
            
            '�v���O���X�o�[�X�V
            objpb.SetTitle "MQ�J���[�ŃZ���̐F��ύX���@" & i & "�F" & Selection.Count
            objpb.SetProgress i / Selection.Count
            
        Next c
        
        '�v���O���X�o�[�̔j��
        Set objpb = Nothing
        
        MsgBox "�������������܂����B"
    End If
    
    '�X�e�[�^�X�o�[�����Z�b�g
    Application.StatusBar = False
    
End Sub

