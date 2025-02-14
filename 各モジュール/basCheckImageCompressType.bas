Attribute VB_Name = "basCheckImageCompressType"
Option Explicit

Sub Check_Image_Compress_Type_Powered_By_ImageMagick(ByVal control As IRibbonControl)
    Dim c As Range
    Dim i As Long
    Dim img As Object
    Dim objpb
    
    '進捗管理のために、0で初期化
    i = 0
    
    If MsgBox("選択セルの右隣のセルに圧縮形式を格納します。" & vbCrLf & "よろしいですか。", vbYesNo) = vbYes Then
        '選択範囲が1列の場合のみ処理をする
        If Selection.Columns.Count = 1 Then
    
            '画像以外のものが出てきてもエラーで止まらないようにする
            On Error Resume Next
            
            ' 進捗バークラスの定義
            Set objpb = New ProgressBar
            
            ' 進捗バーのタイトルを設定
            objpb.SetTitle "圧縮形式確認中"
        
            'ImageMagickを使うためにオブジェクトを作成
            Set img = New ImageMagickObject.MagickImage
            
            For Each c In Selection
                '該当セルが空白の場合は何もしない
                If c.Value <> "" Then
                    c.Offset(0, 1).Value = Split(img.Identify("-format", "%C,", c.Value), ",")(0)
                End If
                
                '進捗のカウントアップ
                i = i + 1
                
                'プログレスバー更新
                objpb.SetTitle "圧縮形式確認中　" & i & "：" & Selection.Count
                objpb.SetProgress i / Selection.Count
                
            Next
            
            Set img = Nothing
            
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



