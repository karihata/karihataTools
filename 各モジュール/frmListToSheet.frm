VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListToSheet 
   Caption         =   "一覧からシートを作成"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   OleObjectBlob   =   "frmListToSheet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmListToSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If MsgBox("選択したセルの値でシートを作成します。" & vbCrLf & "よろしいですか？", vbYesNo) = vbYes Then
         Call ListToSheet(WorkSheetNameCmb.Value)
         Unload frmListToSheet
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    'コンボボックスに項目を追加
    For Each ws In ActiveWorkbook.Worksheets
        WorkSheetNameCmb.AddItem ws.Name
    Next ws
    
    '初期値を指定
    WorkSheetNameCmb.ListIndex = 0
    
    'オブジェクトを解放
    If Not (ws Is Nothing) Then Set ws = Nothing
End Sub
