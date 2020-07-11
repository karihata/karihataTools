Attribute VB_Name = "basListToSheet"
Option Explicit

Sub Show_frmListToSheet(ByVal control As IRibbonControl)
    frmListToSheet.Show
End Sub


Sub ListToSheet(Sheet_Name As String)
    Dim c As Range
    Dim WorkSheets_Dic As New Scripting.Dictionary
    Dim Selection_Dic As New Scripting.Dictionary
    Dim ws As Worksheet
    Dim Skip_Count As Long
    Dim Msg_Str As String
    Dim Sheet_Name_Error As Boolean
    Dim Named_Sheet_Name As String
    
    Skip_Count = 0
    
    'シート作成時にワークシートとのマッチングを行うために既存ワークシートを確保
    For Each ws In ActiveWorkbook.Worksheets
        WorkSheets_Dic.Add ws.Name, 0
    Next ws
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '選択範囲で空白のセル、もしくは選択範囲で重複している場合は何もしない
    For Each c In Selection
        
        'シートネームを文字列で統一するために文字列型の変数に格納
        Named_Sheet_Name = c.Value
        
        If Named_Sheet_Name <> "" And Not (Selection_Dic.Exists(Named_Sheet_Name)) Then
            
            '選択範囲で重複がないか確認するために選択範囲の値を確保
            Selection_Dic.Add Named_Sheet_Name, Named_Sheet_Name
            
            '既存ワークシートとマッチング
            If WorkSheets_Dic.Exists(Named_Sheet_Name) Then
                '処理完了後にスキップ数を表示するためにカウントする
                Skip_Count = Skip_Count + 1
            Else
                'アンマッチの場合
                '指定したワークシートを末尾にコピーしてシート名を変更する
                Worksheets(Sheet_Name).Copy after:=Worksheets(Worksheets.Count)
                
                On Error Resume Next
                
                Worksheets(Worksheets.Count).Name = Named_Sheet_Name
                
                'シート名の変更に失敗した場合、ゴミシートができてしまうため、ゴミシートを削除する
                If Worksheets(Worksheets.Count).Name <> Named_Sheet_Name Then
                    Worksheets(Worksheets.Count).Delete
                    Sheet_Name_Error = True
                End If
                
            End If
            
        End If
    Next c
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'オブジェクトを解放
    If Not (WorkSheets_Dic Is Nothing) Then Set WorkSheets_Dic = Nothing
    If Not (Selection_Dic Is Nothing) Then Set Selection_Dic = Nothing
    If Not (ws Is Nothing) Then Set ws = Nothing
    
    
    Msg_Str = "処理が完了しました。"
    
    '再作成フラグがFALSEでスキップ回数が0より多い場合はスキップを表示する
    If Skip_Count > 0 Then
         Msg_Str = Msg_Str & vbCrLf & "エラー１：既に同じシート名がありましたので、そのシート作成はスキップしました。"
    End If
    
    'シート名変更に失敗したものがある胸をメッセージに追加する
    If Sheet_Name_Error Then
         Msg_Str = Msg_Str & vbCrLf & "エラー２：シート名として設定できないものがありましたので、そのシート作成はスキップしました。"
    End If
    
    'メッセージを表示
    MsgBox Msg_Str
        
    Worksheets(Sheet_Name).Activate
    
End Sub
