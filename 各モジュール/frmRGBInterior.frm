VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRGBInterior 
   Caption         =   "�Z������RGB�l�ŐF��ύX"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5430
   OleObjectBlob   =   "frmRGBInterior.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmRGBInterior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub InteriorChange_Change()
    '�Z���̔w�i�F��ύX�ꍇ�́A�I�v�V�����Ƃ��ĕ����F�𔒂ɂ��邩�`�F�b�N�𐧌䂷��
    If InteriorChange.Value Then
        FontColorWhitechbx.Enabled = True
    Else
        FontColorWhitechbx.Enabled = False
    End If
End Sub

Private Sub SplitEtc_Change()
    '���̑��̃��W�I�{�^����True�ɂȂ�����e�L�X�g�{�b�N�X��L���ɂ���
    If SplitEtc.Value Then
        SplitEtcString.Enabled = True
        SplitEtcString.BackColor = &H80000005
    Else
        SplitEtcString.Enabled = False
        SplitEtcString.BackColor = &H80000010
    End If
End Sub


Private Sub ChangeColorBtn_Click()
    Dim SplitString As String
    Dim i As Long
    
    '��؂蕶��������
    SplitString = ""
    
    
    '���W�I�{�^�������؂蕶�����擾
    For i = 1 To 6
        If Me.Controls("Split" & i).Value = True Then
            SplitString = Mid(Me.Controls("Split" & i).Caption, 2, 1)
        End If
    Next i
    
    '��؂蕶�������̑��ɂȂ��Ă���ꍇ�̓e�L�X�g�{�b�N�X����擾
    If SplitString = "" Then
        '��؂蕶���ł��̑��̏ꍇ�A�e�L�X�g�{�b�N�X����̏ꍇ�̓G���[�ŏ����𒆎~����
        If SplitEtcString.Value <> "" Then
            SplitString = SplitEtcString.Value
            MsgBox "�u���̑��v��I�����͋�؂蕶������͂��ĉ������B"
            Exit Sub
        End If
    End If
    
    
    '�F�ύX�����s
    Call ChangeColorRGB(SplitString, InteriorChange.Caption, FontColorWhitechbx.Value)
End Sub

