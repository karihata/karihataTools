Attribute VB_Name = "basChangeStringColorRGB"
Option Explicit

Sub Show_RGBInterior(ByVal control As IRibbonControl)
    frmRGBInterior.Show
End Sub

    
Sub ChangeColorRGB(SplitString As String, ChangeMode As String, FontColorChangeToWhite As Boolean)
    Dim c As Range
    Dim i As Long
    Dim R As Long
    Dim G As Long
    Dim B As Long
    Dim inputRGBcolor As String
    Dim RGB_Temp As Variant         'RGB���X�v���b�g�����ۂ̊i�[��
    Dim RGB_Check As Boolean        'RGB�̒l���`�F�b�N���Ė��Ȃ����True�ɂȂ�
    
    
    '�i���Ǘ��̂��߂ɁA0�ŏ�����
    i = 0
    
    If MsgBox("�I��͈͂̐F���Z����RGB�l�ɕύX���܂��B" & vbCrLf & "��낵���ł����B" & vbCrLf & vbCrLf & "�u-�v(�n�C�t��)�̃Z���ɂ͉������܂���B", vbYesNo) = vbYes Then
        ' �i���o�[�N���X�̒�`
        Dim objpb
        Set objpb = New ProgressBar
        
        ' �i���o�[�̃^�C�g����ݒ�
        objpb.SetTitle "RGB�J���[�ŃZ���F�ύX��"
        
        '�I��͈͂�����
        For Each c In Selection
            
            '�f�o�b�O�p�Ɋi�[
            inputRGBcolor = c.Value
            
            On Error Resume Next
            
            '�I���Z���̒l���u�����N�ł͂Ȃ����ƁA"-"�ɂȂ��Ă��Ȃ����Ƃ��m�F
            '�Y�������ꍇ�͉������Ȃ�
            If Not (IsEmpty(c.Value)) And inputRGBcolor <> "-" Then
                
                '�l�`�F�b�N�̌��ʂ�������
                RGB_Check = True
                
                '�Z���l����؂蕶���ŃX�v���b�g����
                RGB_Temp = Split(c.Value, SplitString)
                
                '�X�v���b�g���ėv�f����3�̏ꍇ�̂ݏ����𑱍s
                '����ȊO�͏������Ȃ�
                If UBound(RGB_Temp) = 2 Then
                    
                    '�X�v���b�g�セ�ꂼ�ꂪ������0�`255�͈̔͂̐��l�ɂȂ��Ă��邩�m�F
                    'R
                    If Not (IsNumeric(RGB_Temp(0))) Then RGB_Check = False
                    If RGB_Temp(0) < 0 And RGB_Temp(0) > 255 Then RGB_Check = False
                    
                    'G
                    If Not (IsNumeric(RGB_Temp(1))) Then RGB_Check = False
                    If RGB_Temp(1) < 0 And RGB_Temp(1) > 255 Then RGB_Check = False
                    
                    'B
                    If Not (IsNumeric(RGB_Temp(2))) Then RGB_Check = False
                    If RGB_Temp(2) < 0 And RGB_Temp(2) > 255 Then RGB_Check = False
                    
                    'RGB�̂��ꂼ��̒l�`�F�b�N�Ŗ��Ȃ���ΐF��ύX
                    If RGB_Check Then
                        
                        'RGB�l�����ꂼ��i�[
                        R = RGB_Temp(0)
                        G = RGB_Temp(1)
                        B = RGB_Temp(2)
                        
                        
                        '�F�ύX���[�h�ŕ����F���w�i�F������������
                        If ChangeMode = "�Z���̔w�i�F" Then
                            '�Z���̔w�i�F�ύX
                            c.Interior.Color = RGB(R, G, B)
                            
                            '�u�����F�𔒂ɂ���v��ON�̏ꍇ�͕����F�𔒂ɂ���
                            If FontColorChangeToWhite Then
                                '�����������F�𔒂ɂ���ɂ�������RGB����R���g���X�g����v�Z���Ė��Ȃ���Ε����F�𔒂ɂ���
                                If whiteTextColor(R, G, B) Then c.Font.Color = RGB(255, 255, 255)
                            End If
                            
                        Else
                            '�Z���̕����F�ύX
                            c.Font.Color = RGB(R, G, B)
                        End If
                    End If
                    
                End If
                           
            End If
            
            '�i���̃J�E���g�A�b�v
            i = i + 1
            
            '�v���O���X�o�[�X�V
            objpb.SetTitle "RGB�J���[�ŃZ���F�ύX���@" & i & "�F" & Selection.Count
            objpb.SetProgress i / Selection.Count
                        
        Next c
        
        '�v���O���X�o�[�̔j��
        Set objpb = Nothing
        
        MsgBox "�������������܂����B"
    End If
    
    '�X�e�[�^�X�o�[�����Z�b�g
    Application.StatusBar = False
    
End Sub

Public Function whiteTextColor(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer) As Boolean
    '�Q�l�T�C�g
    '�C�ӂ̔w�i�F�ɑ΂��ēǂ݂₷�������F��I��������@
    'https://katashin.info/2018/12/18/247
    Dim R As Double
    Dim G As Double
    Dim B As Double
    
    Dim Lbg As Double
    
    Dim Lw As Integer
    Dim Lb As Integer
    
    Dim Cw As Double
    Dim Cb As Double
    
    'sRGB �� RGB �ɕϊ����A�w�i�F�̑��΋P�x�����߂�
    R = toRgbItem(red)
    G = toRgbItem(green)
    B = toRgbItem(blue)
    Lbg = 0.2126 * R + 0.7152 * G + 0.0722 * B
    
    
    '���ƍ��̑��΋P�x�B��`���炻�ꂼ�� 1 �� 0 �ɂȂ�B
    Lw = 1
    Lb = 0
    
    '���Ɣw�i�F�̃R���g���X�g��A���Ɣw�i�F�̃R���g���X�g���
    '���ꂼ�ꋁ�߂�B
    Cw = (Lw + 0.05) / (Lbg + 0.05)
    Cb = (Lbg + 0.05) / (Lb + 0.05)
    
    '�R���g���X�g�䂪�傫�����𕶎��F�Ƃ��ĕԂ��B
    If Cw < Cb Then
        whiteTextColor = False
    Else
        whiteTextColor = True
    End If
    
    
End Function

'sRGB �� RGB �ɕϊ����A�w�i�F�̑��΋P�x�����߂�
Public Function toRgbItem(item As Integer) As Double
    Dim i As Double
    
    i = item / 255
    
    If i <= 0.03928 Then
        toRgbItem = i / 12.92
    Else
        toRgbItem = ((i + 0.055) / 1.055) ^ 2.4
    End If
    
End Function

