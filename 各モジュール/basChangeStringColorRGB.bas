Attribute VB_Name = "basChangeStringColorRGB"
Option Explicit

Sub Show_RGBInterior(ByVal control As IRibbonControl)
    frmRGBInterior.Show
End Sub

    
Sub ChangeColorRGB(SplitString As String, ChangeMode As String, FontColorChangeToWhite As Boolean)
    Dim c As Range
    Dim i As Long
    Dim R As Long
    Dim G As Long
    Dim B As Long
    Dim inputRGBcolor As String
    Dim RGB_Temp As Variant         'RGBをスプリットした際の格納先
    Dim RGB_Check As Boolean        'RGBの値をチェックして問題なければTrueになる
    
    
    '進捗管理のために、0で初期化
    i = 0
    
    If MsgBox("選択範囲の色をセルのRGB値に変更します。" & vbCrLf & "よろしいですか。" & vbCrLf & vbCrLf & "「-」(ハイフン)のセルには何もしません。", vbYesNo) = vbYes Then
        ' 進捗バークラスの定義
        Dim objpb
        Set objpb = New ProgressBar
        
        ' 進捗バーのタイトルを設定
        objpb.SetTitle "RGBカラーでセル色変更中"
        
        '選択範囲を巡回
        For Each c In Selection
            
            'デバッグ用に格納
            inputRGBcolor = c.Value
            
            On Error Resume Next
            
            '選択セルの値がブランクではないこと、"-"になっていないことを確認
            '該当した場合は何もしない
            If Not (IsEmpty(c.Value)) And inputRGBcolor <> "-" Then
                
                '値チェックの結果を初期化
                RGB_Check = True
                
                'セル値を区切り文字でスプリットする
                RGB_Temp = Split(c.Value, SplitString)
                
                'スプリットして要素数が3つの場合のみ処理を続行
                'それ以外は処理しない
                If UBound(RGB_Temp) = 2 Then
                    
                    'スプリット後それぞれが整数で0～255の範囲の数値になっているか確認
                    'R
                    If Not (IsNumeric(RGB_Temp(0))) Then RGB_Check = False
                    If RGB_Temp(0) < 0 And RGB_Temp(0) > 255 Then RGB_Check = False
                    
                    'G
                    If Not (IsNumeric(RGB_Temp(1))) Then RGB_Check = False
                    If RGB_Temp(1) < 0 And RGB_Temp(1) > 255 Then RGB_Check = False
                    
                    'B
                    If Not (IsNumeric(RGB_Temp(2))) Then RGB_Check = False
                    If RGB_Temp(2) < 0 And RGB_Temp(2) > 255 Then RGB_Check = False
                    
                    'RGBのそれぞれの値チェックで問題なければ色を変更
                    If RGB_Check Then
                        
                        'RGB値をそれぞれ格納
                        R = RGB_Temp(0)
                        G = RGB_Temp(1)
                        B = RGB_Temp(2)
                        
                        
                        '色変更モードで文字色か背景色かを処理分岐
                        If ChangeMode = "セルの背景色" Then
                            'セルの背景色変更
                            c.Interior.Color = RGB(R, G, B)
                            
                            '「文字色を白にする」がONの場合は文字色を白にする
                            If FontColorChangeToWhite Then
                                'ただし文字色を白にするにあたってRGBからコントラスト比を計算して問題なければ文字色を白にする
                                If whiteTextColor(R, G, B) Then c.Font.Color = RGB(255, 255, 255)
                            End If
                            
                        Else
                            'セルの文字色変更
                            c.Font.Color = RGB(R, G, B)
                        End If
                    End If
                    
                End If
                           
            End If
            
            '進捗のカウントアップ
            i = i + 1
            
            'プログレスバー更新
            objpb.SetTitle "RGBカラーでセル色変更中　" & i & "：" & Selection.Count
            objpb.SetProgress i / Selection.Count
                        
        Next c
        
        'プログレスバーの破棄
        Set objpb = Nothing
        
        MsgBox "処理が完了しました。"
    End If
    
    'ステータスバーをリセット
    Application.StatusBar = False
    
End Sub

Public Function whiteTextColor(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer) As Boolean
    '参考サイト
    '任意の背景色に対して読みやすい文字色を選択する方法
    'https://katashin.info/2018/12/18/247
    Dim R As Double
    Dim G As Double
    Dim B As Double
    
    Dim Lbg As Double
    
    Dim Lw As Integer
    Dim Lb As Integer
    
    Dim Cw As Double
    Dim Cb As Double
    
    'sRGB を RGB に変換し、背景色の相対輝度を求める
    R = toRgbItem(red)
    G = toRgbItem(green)
    B = toRgbItem(blue)
    Lbg = 0.2126 * R + 0.7152 * G + 0.0722 * B
    
    
    '白と黒の相対輝度。定義からそれぞれ 1 と 0 になる。
    Lw = 1
    Lb = 0
    
    '白と背景色のコントラスト比、黒と背景色のコントラスト比を
    'それぞれ求める。
    Cw = (Lw + 0.05) / (Lbg + 0.05)
    Cb = (Lbg + 0.05) / (Lb + 0.05)
    
    'コントラスト比が大きい方を文字色として返す。
    If Cw < Cb Then
        whiteTextColor = False
    Else
        whiteTextColor = True
    End If
    
    
End Function

'sRGB を RGB に変換し、背景色の相対輝度を求める
Public Function toRgbItem(item As Integer) As Double
    Dim i As Double
    
    i = item / 255
    
    If i <= 0.03928 Then
        toRgbItem = i / 12.92
    Else
        toRgbItem = ((i + 0.055) / 1.055) ^ 2.4
    End If
    
End Function

