Attribute VB_Name = "basPassCheck"
Option Explicit

Sub Path_Check(ByVal control As IRibbonControl)
    Dim c As Range
    Dim i As Long
    Dim objpb
    
    '進捗管理のために、0で初期化
    i = 0
    
    If MsgBox("選択セルの右隣のセルに結果を格納します。" & vbCrLf & "よろしいですか。", vbYesNo) = vbYes Then
        '選択範囲が1列の場合のみ処理をする
        If Selection.Columns.Count = 1 Then
            
            ' 進捗バークラスの定義
            Set objpb = New ProgressBar
            
            ' 進捗バーのタイトルを設定
            objpb.SetTitle "パス確認"
            
            '選択範囲を巡回
            For Each c In Selection
                
                On Error Resume Next
                
                '選択セルがブランク以外の場合は処理を続ける
                If c.Value <> "" Then
                    c.Offset(0, 1).Value = FolderExists(c.Value)
                Else
                    '選択セルがブランクの場合は何もしない
                    c.Offset(0, 1).Value = ""
                End If
                
                '進捗のカウントアップ
                i = i + 1
                
                'プログレスバー更新
                objpb.SetTitle "パス確認　" & i & "：" & Selection.Count
                objpb.SetProgress i / Selection.Count
            Next c
            
            'プログレスバーの破棄
            Set objpb = Nothing
            
            MsgBox "処理が完了しました。"
        Else
            '選択範囲が1列ではなかった場合はエラーメッセージを表示
            MsgBox "1列だけ選択してください。"
        End If
    End If
    
    'ステータスバーをリセット
    Application.StatusBar = False
    
End Sub

Function FolderExists(Path_String As String) As Boolean
  'FileSystemObjectでファイル存在チェックをする
  Dim fso As New Scripting.FileSystemObject
  
  '何らかのエラーが発生した場合はErrorTrapの処理を実行
  On Error GoTo ErrorTrap
  
  If fso.FileExists(Path_String) Then
    FolderExists = True
  Else
    FolderExists = False
  End If
    
  'オブジェクトの片付け
  Set fso = Nothing
  
Exit Function

ErrorTrap:
    If Not (fso Is Nothing) Then Set fso = Nothing
    FolderExists = False
End Function
