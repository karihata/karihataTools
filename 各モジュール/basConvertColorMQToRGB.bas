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
    
    '進捗管理のために、0で初期化
    i = 0
    
    If MsgBox("選択範囲をセルのMQカラー値で塗りつぶします。" & vbCrLf & "よろしいですか。" & vbCrLf & vbCrLf & "「-」(ハイフン)のセルには何もしません。", vbYesNo) = vbYes Then
        ' 進捗バークラスの定義
        Set objpb = New ProgressBar
        
        ' 進捗バーのタイトルを設定
        objpb.SetTitle "MQカラーでセルの色を変更中"
        
        '選択範囲を巡回
        For Each c In Selection
            
            'デバッグ用に格納
            inputMQcolor = c.Value
            
            On Error Resume Next
            
            '選択セルの値がブランクではないこと、"-"になっていないことを確認
            '該当した場合は何もしない
            If Not (IsEmpty(c.Value)) And inputMQcolor <> "-" Then
                
                '選択セルの値の整数チェック
                '整数ではない場合は何もしない
                If IsNumeric(inputMQcolor) And c.Value = Int(inputMQcolor) Then
                    'セルの値からMQカラー番号からRGBの番号にそれぞれ変換して格納
                    '256 * 256 = 65536
                    B = inputMQcolor \ 65536
                    
                    G = (inputMQcolor - (B * 65536)) \ 256
                    
                    R = inputMQcolor - (G * 256) - (B * 65536)
                                  
                    'セルの色変更
                    c.Interior.Color = RGB(R, G, B)
                End If
            
            End If
            
            '進捗のカウントアップ
            i = i + 1
            
            'プログレスバー更新
            objpb.SetTitle "MQカラーでセルの色を変更中　" & i & "：" & Selection.Count
            objpb.SetProgress i / Selection.Count
            
        Next c
        
        'プログレスバーの破棄
        Set objpb = Nothing
        
        MsgBox "処理が完了しました。"
    End If
    
    'ステータスバーをリセット
    Application.StatusBar = False
    
End Sub

