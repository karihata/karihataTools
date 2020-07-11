VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImageCount 
   Caption         =   "画像のステータスチェック"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3615
   OleObjectBlob   =   "frmImageCount.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmImageCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    '複数のファイルパス
    Dim Image_Files As Variant
    
    '画像ファイルをダイアログから選択する
    Image_Files = Application.GetOpenFilename(FileFilter:="すべてのファイル(*.*),*.*", Title:="画像のステータスチェック", MultiSelect:=True)
    
    'ダイアログでキャンセルもしくは閉じられた場合は処理を中止
    If IsArray(Image_Files) <> False Then
        Call Check_Image_Type_Powered_By_ImageMagick(Image_Files, chkPageCount, chkCompress, chkPixelCount, chkDPI)
    End If
    Unload frmImageCount
End Sub

Private Sub CommandButton2_Click()
    'メッセージを表示してからスタートする
    If MsgBox("選択セルの値から画像のステータスチェックをします。" & vbCrLf & "よろしいですか。", vbYesNo) = vbYes Then
        '複数のファイルパス
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
