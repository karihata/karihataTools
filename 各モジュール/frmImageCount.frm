VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImageCount 
   Caption         =   "�摜�̃X�e�[�^�X�`�F�b�N"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3615
   OleObjectBlob   =   "frmImageCount.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmImageCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    '�����̃t�@�C���p�X
    Dim Image_Files As Variant
    
    '�摜�t�@�C�����_�C�A���O����I������
    Image_Files = Application.GetOpenFilename(FileFilter:="���ׂẴt�@�C��(*.*),*.*", Title:="�摜�̃X�e�[�^�X�`�F�b�N", MultiSelect:=True)
    
    '�_�C�A���O�ŃL�����Z���������͕���ꂽ�ꍇ�͏����𒆎~
    If IsArray(Image_Files) <> False Then
        Call Check_Image_Type_Powered_By_ImageMagick(Image_Files, chkPageCount, chkCompress, chkPixelCount, chkDPI)
    End If
    Unload frmImageCount
End Sub

Private Sub CommandButton2_Click()
    '���b�Z�[�W��\�����Ă���X�^�[�g����
    If MsgBox("�I���Z���̒l����摜�̃X�e�[�^�X�`�F�b�N�����܂��B" & vbCrLf & "��낵���ł����B", vbYesNo) = vbYes Then
        '�����̃t�@�C���p�X
        Dim Image_Files As Variant
        Dim c
        
        Image_Files = Array()
        
        For Each c In Selection
            Call Push(Image_Files, c.Value)
        Next
        
        Call Check_Image_Type_Powered_By_ImageMagick(Image_Files, chkPageCount, chkCompress, chkPixelCount, chkDPI)
    End If
    
    Unload frmImageCount
End Sub
