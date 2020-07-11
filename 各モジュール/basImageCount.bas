Attribute VB_Name = "basImageCount"
Option Explicit

Sub frmImageCount_Show(ByVal control As IRibbonControl)
    frmImageCount.Show
End Sub

Sub Check_Image_Type_Powered_By_ImageMagick(Image_Files As Variant, Page_flg As Boolean, Compress_flg As Boolean, Pixel_flg As Boolean, DPI_flg As Boolean)
    Dim Image_File As Variant                       '単体のファイルパス
    Dim File_Count As Long                          '進捗表示用の選択ファイル数を格納
    Dim i As Long                                   'ファイルごとのインデックス
    Dim j As Long                                   '項目ごとのインデックス
    Dim img As Object                               'ImageMagickのオブジェクト
    Dim Headder_Dic As New Scripting.Dictionary     'ヘッダー項目の列番地格納用
    Dim c As Variant                                '項目のkey格納
    Dim objpb                                       '進捗バーのオブジェクト
    
    '描画更新しない
    Application.ScreenUpdating = False
    
    '画像以外のものが出てきてもエラーで止まらないようにする
    On Error Resume Next
    
    '進捗表示用のファイル選択数を格納
    File_Count = UBound(Image_Files)
    
    'ImageMagickを使うためにオブジェクトを作成
    Set img = New ImageMagickObject.MagickImage
    
    'シートを新規作成して一番後ろに配置
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    
    '展開するデータの開始行
    i = 1
    
    '項目インデックスの初期化
    j = 1
    
    '各チェックボックスのフラグから連想配列に変換する
    
    'これだけは固定
    Headder_Dic.Add "ファイルパス", j
    
    'ここからフラグから連想配列に変換
    If Page_flg Then
        Headder_Dic.Add "ページ数", j + 1
        j = j + 1
    End If
    
    If Compress_flg Then
        Headder_Dic.Add "圧縮形式", j + 1
        j = j + 1
    End If
    
    If Pixel_flg Then
        Headder_Dic.Add "ピクセル数_縦", j + 1
        Headder_Dic.Add "ピクセル数_横", j + 2
        j = j + 2
    End If
    
    If DPI_flg Then
        Headder_Dic.Add "DPI_縦", j + 1
        Headder_Dic.Add "DPI_横", j + 2
        j = j + 2
    End If
        
    'ヘッダーをセルに格納
    For Each c In Headder_Dic.Keys
        Cells(1, Headder_Dic(c)).Value = c
    Next
    
    '進捗バークラスの定義
    Set objpb = New ProgressBar
    
    '進捗バーのタイトルを設定
    objpb.SetTitle "画像のステータスチェック"
    
    
    '各ファイルを開いていく
    For Each Image_File In Image_Files
        '該当セルが空白の場合は何もしない
        If Image_File <> "" Then
            
            'ファイルパス
            Cells(i + 1, Headder_Dic("ファイルパス")).Value = Image_File
            
            'ページ数
            If Page_flg Then Cells(i + 1, Headder_Dic("ページ数")).Value = UBound(Split(img.Identify("-format", "%p,", Image_File), ","))
            
            '圧縮形式
            If Compress_flg Then Cells(i + 1, Headder_Dic("圧縮形式")).Value = Split(img.Identify("-format", "%C,", Image_File), ",")(0)
            
            'ピクセル数
            If Pixel_flg Then
                Cells(i + 1, Headder_Dic("ピクセル数_縦")).Value = Split(img.Identify("-format", "%h,", Image_File), ",")(0)
                Cells(i + 1, Headder_Dic("ピクセル数_横")).Value = Split(img.Identify("-format", "%w,", Image_File), ",")(0)
            End If
            
            'DPI
            If DPI_flg Then
                Cells(i + 1, Headder_Dic("DPI_縦")).Value = Split(img.Identify("-format", "%y,", Image_File), ",")(0)
                Cells(i + 1, Headder_Dic("DPI_横")).Value = Split(img.Identify("-format", "%x,", Image_File), ",")(0)
            End If
            
        End If
        
        'プログレスバー更新
        objpb.SetTitle "画像のステータスチェック　" & i & "：" & (File_Count + 1)
        objpb.SetProgress i / (File_Count + 1)
        
        '進捗のカウントアップ
        i = i + 1
    Next
    
    Set img = Nothing
    
    'プログレスバーの破棄
    Set objpb = Nothing
    
    'ヘッダー文字の幅で列幅を自動調整
    Range("A1:G1").Columns.AutoFit
    
    'ステータスバーをリセット
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox ("処理が完了しました")
    
End Sub
