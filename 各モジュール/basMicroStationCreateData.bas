Attribute VB_Name = "basMicroStationCreateData"
Option Explicit

'選択範囲内の「点名とXとY」をマイクロステーションのXYZインポートテキストを出力する
'必ず点名、X座標、Y座標の列の順番になっていること
Sub CreateXYText(ByVal control As IRibbonControl)
    If MsgBox("XYテキストファイルを出力します。" & vbCrLf & "選択範囲には「点名」「X座標」「Y座標」の順番になっていますか？" & vbCrLf & "※先頭行は見出しではなくデータがある行にして下さい。" _
    , vbYesNo + vbQuestion, "確認") = vbYes Then
        
        '選択範囲で3列選択していない場合はエラーメッセージを表示して処理をしない
        If Selection.Columns.Count = 3 Then
            Dim Name_txt_File As Variant        '出力先の点名txtファイルパス
            Dim Point_txt_File As Variant       '出力先の点txtファイルパス
            Dim i As Long
            
            '初期化
            Name_txt_File = False
            Point_txt_File = False
            
            '何らかのエラーが発生した場合はErrorTrapの処理を実行
            On Error GoTo ErrorTrap
            
            '保存ファイル名をそれぞれ決めてもらう
            Name_txt_File = Application.GetSaveAsFilename(InitialFileName:=ActiveSheet.Name & "_PointName.txt", FileFilter:="txtファイル(*.txt),*.txt", Title:="点名txtファイルの保存先")
            Point_txt_File = Application.GetSaveAsFilename(InitialFileName:=ActiveSheet.Name & "_Point.txt", FileFilter:="txtファイル(*.txt),*.txt", Title:="点txtファイルの保存先")
            
            
            '保存ファイル名のどちらかがFalseの場合は処理を行わない
            If Name_txt_File <> False And Point_txt_File <> False Then
                
                '点名txtファイルをオープン
                Open Name_txt_File For Output As #1
                    
                    '選択した行の分だけテキストファイルに書き込む
                    For i = 0 To Selection.Rows.Count - 1
                        '点情報を書き込む
                        '書き込む際に点名に全角半角スペースが含まれている場合はアンダースコアに置換える
                        Print #1, Replace(Replace(Selection(1).Offset(i, 0).Value, " ", "_"), "　", "_") & " " & Selection(1).Offset(i, 1).Value & " " & Selection(1).Offset(i, 2).Value & " " & "0" & vbCrLf;
                    Next i
                    
                Close #1
                
                Open Point_txt_File For Output As #1
                    
                    '選択した行の分だけテキストファイルに書き込む
                    For i = 0 To Selection.Rows.Count - 1
                        '点情報を書き込む
                        Print #1, Selection(1).Offset(i, 1).Value & " " & Selection(1).Offset(i, 2).Value & " " & "0" & vbCrLf;
                    Next i
                    
                Close #1
                
                MsgBox "テキストファイルを出力しました。"
            End If

        Else
            MsgBox "「点名」「X座標」「Y座標」の3列を選択して下さい。"
        
        End If
        
        

    End If
    
Exit Sub
ErrorTrap:
    MsgBox "エラーが発生しました。" & vbCrLf & "エラー番号：" & Err.Number & vbCrLf & "エラー内容" & vbCrLf & Err.Description
    Close #1
End Sub
