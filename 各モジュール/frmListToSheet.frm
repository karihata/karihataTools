VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListToSheet 
   Caption         =   "�ꗗ����V�[�g���쐬"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   OleObjectBlob   =   "frmListToSheet.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmListToSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If MsgBox("�I�������Z���̒l�ŃV�[�g���쐬���܂��B" & vbCrLf & "��낵���ł����H", vbYesNo) = vbYes Then
         Call ListToSheet(WorkSheetNameCmb.Value)
         Unload frmListToSheet
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    '�R���{�{�b�N�X�ɍ��ڂ�ǉ�
    For Each ws In ActiveWorkbook.Worksheets
        WorkSheetNameCmb.AddItem ws.Name
    Next ws
    
    '�����l���w��
    WorkSheetNameCmb.ListIndex = 0
    
    '�I�u�W�F�N�g�����
    If Not (ws Is Nothing) Then Set ws = Nothing
End Sub
