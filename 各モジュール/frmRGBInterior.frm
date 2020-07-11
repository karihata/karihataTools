VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRGBInterior 
   Caption         =   "セル内のRGB値で色を変更"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5430
   OleObjectBlob   =   "frmRGBInterior.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmRGBInterior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub InteriorChange_Change()
    'セルの背景色を変更場合は、オプションとして文字色を白にするかチェックを制御する
    If InteriorChange.Value Then
        FontColorWhitechbx.Enabled = True
    Else
        FontColorWhitechbx.Enabled = False
    End If
End Sub

Private Sub SplitEtc_Change()
    'その他のラジオボタンがTrueになったらテキストボックスを有効にする
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
    
    '区切り文字初期化
    SplitString = ""
    
    
    'ラジオボタンから区切り文字を取得
    For i = 1 To 6
        If Me.Controls("Split" & i).Value = True Then
            SplitString = Mid(Me.Controls("Split" & i).Caption, 2, 1)
        End If
    Next i
    
    '区切り文字がその他になっている場合はテキストボックスから取得
    If SplitString = "" Then
        '区切り文字でその他の場合、テキストボックスが空の場合はエラーで処理を中止する
        If SplitEtcString.Value <> "" Then
            SplitString = SplitEtcString.Value
            MsgBox "「その他」を選択時は区切り文字を入力して下さい。"
            Exit Sub
        End If
    End If
    
    
    '色変更を実行
    Call ChangeColorRGB(SplitString, InteriorChange.Caption, FontColorWhitechbx.Value)
End Sub

