Attribute VB_Name = "NewEraFunction"
Option Explicit

'#########################################################################
'#
'#    [新元号:令和]対応 日付変換関数  Ver 1.20
'#
'#       EraFormat ( シリアル値 ⇒ 日付文字列    )
'#       EraCDate  ( 日付文字列 ⇒ シリアル値    )
'#       EraIsDate ( 日付データ ⇒ True or False )
'#
'#       Ver 0.10 , 2018/12/ 1  暫定版 初版 ( EraFormat / EraCDate )
'#       Ver 0.20 , 2019/ 1/ 3  暫定版 ２版
'#
'#       Ver 1.00 , 2019/4/3
'#         (1) 正式版(EraFormat / EraCDate)リリース
'#
'#       Ver 1.10 , 2019/4/9
'#         (1) EraIsDate を追加
'#             それに伴い[EraCDateのセル制限(0～60)]を解除
'#
'#       Ver 1.20 , 2019/4/13
'#         (1) EraCDateの[日付]引数で[数値]をサポートします(シリアル値と見做します)
'#         (2) EraCDateの[日付]引数で[日付文字列＋時刻文字列]をサポートします
'#             EraFormat/EraIsDateの[日付]引数でも同様にサポートします
'#         (3) EraCDateの[日付]引数で西暦年3桁(100～999年),
'#             西暦年2桁(2000年代と解釈)をサポートします
'#         (4) [令和]未対応環境でもEraFormatの[日付]引数に
'#             [令和]日付文字列を指定可能とします
'#
'#    作者: AddinBox 角田 桂一
'#          ( http://addinbox.sakura.ne.jp/Excel_Tips28.htm )
'#
'#    -- 使用条件 --
'#    (a) EraFormat/EraCDate/EraIsDate 関数はフリーウェアです。
'#        御自由に各自のプログラムに組み込んで利用して戴いて構いません｡
'#        但し､プログラム先頭のコメントも必ず一緒に コピーしてください｡
'#
'#    (b) EraFormat/EraCDate/EraIsDate 関数を組み込んだプログラムの
'#        再頒布にも制限はありません。
'#
'#########################################################################



'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/    [新元号:令和]対応 日付変換関数 EraFormat ( シリアル値 ⇒ 日付文字列 )
'_/
'_/    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'_/    新元号に対応していないシステム(Office2007以前 or 対応
'_/    アップデートを施していないOffice2010以降)でも、新元号に
'_/    基づく和暦変換を可能にする関数です。
'_/
'_/    ExcelのTEXT関数/VBAのFormat関数の代わりに使用してください。
'_/    [元号/和暦年]以外の編集文字を一緒に使用しても問題ありません。
'_/
'_/    また、新元号に対応済みのシステムで使用しても問題ありません。
'_/
'_/    尚、EraFormat は AddinBox/kt関数アドイン(Ver5.30)に
'_/    ktEraFormat の名前で収録します。
'_/    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'_/    (注) 日付にシリアル値[0～60]の値を指定した場合に得られる日付編集は、
'_/         シート上の書式/TEXT関数で得られる結果と１日ズレます。
'_/      EraFormat: 0⇒1899/12/30, 1⇒1899/12/31, 2⇒1900/1/1, 60⇒1900/2/28, 61⇒1900/3/1
'_/      シート上 : 0⇒1900/1/0  , 1⇒1900/1/1  , 2⇒1900/1/2, 60⇒1900/2/29, 61⇒1900/3/1
'_/      (Excelは Lotus1-2-3互換の為に[1900/1/1～1900/2/29]のシリアル値を敢えてズラしています)
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Function EraFormat(ByVal 日付 As Variant, ByVal 編集文字 As String, _
                          Optional ByVal 元年表記 As Boolean = False) As Variant

Const cst新元号開始日 As Date = #5/1/2019#

Const cstEra1 As String = "令和"
Const cstEra2 As String = "令"
Const cstEra3 As String = "R"

Const cstReplace_Era3 As String = "α"  '和暦年の迂回置換用の代用文字
Const cstReplace_ee As String = "β"
Const cstReplace_e As String = "γ"
Const cstReplace_Period As String = "Ω"    'ピリオド区切りの代用文字

Dim dtm日付 As Date
Dim strDateFormat As String
Dim Result As String

    '(注意事項)
    ' 1. LCase/UCase による小文字/大文字への統一変換は、
    '    g/e以外の編集定義を壊すかもしれないので使わない
    ' 2. 元号半角 および 和暦年編集後の "0" が数値/日付編集文字と
    '    重複するケースを回避するために、Format関数の実施前では
    '    代用文字(α,β,γ)で迂回置換し、Format実施後に改めて置換する。
    ' 3. 正規表現による検索/置換は利用しません(VBAではVBScriptを必要とする為)

    If EraIsDate(日付) Then
        '[令和]未対応環境でも [日付]引数に [令和]日付文字列を指定可能とする
        '[元年]表記の日付文字列も可としているので、[Formatで対応可能]判定の前に
        'EraCDateでシリアル値変換を行なう
        dtm日付 = EraCDate(日付)
    Else
        EraFormat = CVErr(xlErrValue)
        Exit Function
    End If

    '--- rr書式 ( = gggee ) / r書式 ( = ee)の変換(Format関数では不可) ---
    ' (rr ⇒ r の順で置換する)
    ' ※ LCase/UCase非使用の注意事項を参照
    strDateFormat = Replace(編集文字, "rr", "gggee")
    strDateFormat = Replace(strDateFormat, "rR", "gggee")
    strDateFormat = Replace(strDateFormat, "Rr", "gggee")
    strDateFormat = Replace(strDateFormat, "RR", "gggee")
    
    strDateFormat = Replace(strDateFormat, "r", "ee")
    strDateFormat = Replace(strDateFormat, "R", "ee")



    '以下の何れかの条件では問題がないので全て Format関数に任せて完了とする。
    ' (1) [平成]以前(2019/4/30 以前)の日付
    ' (2) 新元号対応バージョン or 新元号対応アップデート実施済の環境
    If (dtm日付 < cst新元号開始日) Or _
       (Format(cst新元号開始日, "geemmdd") = "R010501") Then
        On Error Resume Next
        Result = Format(dtm日付, strDateFormat)    ' Format関数による編集
        If (Err.Number <> 0) Then
            EraFormat = CVErr(xlErrValue)
            Exit Function
        End If
        On Error GoTo 0
        ' 元年⇔1年 編集
        EraFormat = prvFirstYearEdit(元年表記, dtm日付, Result, 編集文字)
        Exit Function
    End If



    '####################################################################
    '###                                                              ###
    '###  以降 [日付≧2019/5/1 ＆ [新元号]未対応環境] 限定の変換処理  ###
    '###                                                              ###
    '####################################################################

    'ロケールIDが有れば取り除く([$-411] : 日本 , [$-409] : 米国)
    'Format関数ではロケールIDが無くても期待通りに変換される
    strDateFormat = Replace(strDateFormat, "[$-411]", "")
    strDateFormat = Replace(strDateFormat, "[$-409]", "")
    If (strDateFormat = "") Then
        EraFormat = CVErr(xlErrValue)
        Exit Function
    End If
    
    '--- ggg⇒"令和" , gg⇒"令" , g⇒"R"(代用文字α) に変換する ---
    ' (ggg ⇒ gg ⇒ g の順で置換する)
    ' ※ LCase/UCase非使用の注意事項を参照
    strDateFormat = Replace(strDateFormat, "ggg", cstEra1)
    strDateFormat = Replace(strDateFormat, "Ggg", cstEra1)
    strDateFormat = Replace(strDateFormat, "gGg", cstEra1)
    strDateFormat = Replace(strDateFormat, "ggG", cstEra1)
    strDateFormat = Replace(strDateFormat, "GGg", cstEra1)
    strDateFormat = Replace(strDateFormat, "GgG", cstEra1)
    strDateFormat = Replace(strDateFormat, "gGG", cstEra1)
    strDateFormat = Replace(strDateFormat, "GGG", cstEra1)
    
    strDateFormat = Replace(strDateFormat, "gg", cstEra2)
    strDateFormat = Replace(strDateFormat, "gG", cstEra2)
    strDateFormat = Replace(strDateFormat, "Gg", cstEra2)
    strDateFormat = Replace(strDateFormat, "GG", cstEra2)
    
    strDateFormat = Replace(strDateFormat, "g", cstReplace_Era3)
    strDateFormat = Replace(strDateFormat, "G", cstReplace_Era3)
    
    '--- ee / e を和暦年(西暦年 - 2018) に変換する(代用文字 β,γ) ---
    ' ※ ここまで処理が流れて来るのは2019/5/1以降の日付のみ。
    '    平成以前(2019/4/30以前)の日付は、ここまで流れて来ないので
    '    和暦年の換算式は[西暦年-2018]固定で大丈夫
    ' (補) [和暦年]編集文字は e / ee のみ。eee は無い。
    '      eee年は [ee]+[e年]と解釈(1年⇒ 011年 になる)される。
    
    ' (ee ⇒ e の順で置換する)
    strDateFormat = Replace(strDateFormat, "ee", cstReplace_ee)
    strDateFormat = Replace(strDateFormat, "eE", cstReplace_ee)
    strDateFormat = Replace(strDateFormat, "Ee", cstReplace_ee)
    strDateFormat = Replace(strDateFormat, "EE", cstReplace_ee)
    
    strDateFormat = Replace(strDateFormat, "e", cstReplace_e)
    strDateFormat = Replace(strDateFormat, "E", cstReplace_e)
    
    
    '---【 g , e 以外の編集は Format関数に任せる 】---
    '
    ' 但し、Formmat に任せる前に ピリオドも代用文字に置換しておく必要がある。
    ' 理由："ge.m.d"⇒"αγ.m.d" となるが、[年]編集文字が無くなる為に
    '       最初のピリオドが小数点と解釈されてしまいます。
    '       その結果、シリアル値がピリオド位置に数値として表示されます。
    '       残りの "m.d" も月日編集文字ではなく単なる固定表示文字として
    '       解釈されて、m.d の文字でそのまま表示されます。
    strDateFormat = Replace(strDateFormat, ".", cstReplace_Period)
    
    On Error Resume Next
    Result = Format(dtm日付, strDateFormat)    ' Format関数による編集
    If (Err.Number <> 0) Then
        EraFormat = CVErr(xlErrValue)
        Exit Function
    End If
    On Error GoTo 0
    '----------------------------------------------------

    ' 迂回置換の代用文字(α,β,γ,Ω)に
    ' 本来の値(元号半角, 和暦年2桁, 和暦年1桁, ピリオド)を書き込む
    ' ※ ここまで処理が流れて来るのは2019/5/1以降の日付のみ。
    '    平成以前(2019/4/30以前)の日付は、ここまで流れて来ないので
    '    和暦年の換算式は[西暦年-2018]固定で大丈夫
    Result = Replace(Result, cstReplace_Era3, cstEra3)
    Result = Replace(Result, cstReplace_ee, Format(Year(dtm日付) - 2018, "00"))
    Result = Replace(Result, cstReplace_e, Format(Year(dtm日付) - 2018, "0"))
    Result = Replace(Result, cstReplace_Period, ".")


    ' 元年⇔1年 編集
    EraFormat = prvFirstYearEdit(元年表記, dtm日付, Result, 編集文字)

End Function

'-----------------------------------------------------------------------------
' "平成01年"/"平成1年"/"平01年"/"平1年"/"H01年"/"H1年" 等を"元年"表記に改めます。
'
' (注) 改元関連のアップデートでVBAのFormat関数自体に「元年」表記の機能が追加されました。
'      (レジストリ(InitialEraYear)で[元年]指定が必要)
'                              -- アップデート済環境 , 未アップデート環境
'   Format("2019/5/1","ggge年") ⇒   "令和元年"      ,  "令和1年
'   Format("2019/5/1","gge年")  ⇒   "令元年"        ,  "令1年"
'   Format("2019/5/1","ge年")   ⇒   "R元年"         ,  "R1年"
'
'  Format関数により既に"元年"と編集されているケースでは、
'  EraFormat関数の[元年表記]引数の指定に合わせて、
'  False(元年表記なし)の場合には "元年"⇒"1年" or "01年" に戻します。
'-----------------------------------------------------------------------------
Private Function prvFirstYearEdit(ByVal FirstYear As Boolean, _
                                  ByVal SerialDate As Date, _
                                  ByVal EditDate As String, _
                                  ByVal EditPattern As String) As String
Dim strEdit As String

    If (LCase(EditPattern) Like "*e年*") Then
        ' [和暦+"年"]編集あり
    Else
        ' [和暦+"年"]編集なし … [1年⇔元年]変換処理は不要
        prvFirstYearEdit = EditDate
        Exit Function
    End If

    If (FirstYear = False) Then
        ' "令和元年"⇒"令和1年" 等、"1年"表記に戻す
        ' 既に"元年"表記になっているものなので月日までチェックする必要なし
        Select Case Year(SerialDate)
          Case 1868, 1912, 1926, 1989, 2019
            ' 明治1年(1868), 大正1年(1912), 昭和1年(1926), 平成1年(1989), 令和1年(2019)
            strEdit = prvFormat_GannenTo1st(EditDate, EditPattern)
            
          Case Else
            '明治以前 or 各元号の[２年～]
            strEdit = EditDate
        End Select
    Else
        ' "令和1年"⇒"令和元年" 等、"元年"表記に改める
        Select Case Format(SerialDate, "yyyymmdd")
          Case "18681023" To "18681231"     '明治 元年
            strEdit = prvFormat_1stToGannen(EditDate)
        
          Case "19120730" To "19121231"     '大正 元年
            strEdit = prvFormat_1stToGannen(EditDate)
        
          Case "19261225" To "19261231"     '昭和 元年
            strEdit = prvFormat_1stToGannen(EditDate)
        
          Case "19890108" To "19891231"     '平成 元年
            strEdit = prvFormat_1stToGannen(EditDate)
        
         Case "20190501" To "20191231"     '令和 元年
            strEdit = prvFormat_1stToGannen(EditDate)
        
         Case Else
            '明治以前 or 各元号の[２年～]
            strEdit = EditDate
       End Select
    End If
    
    prvFirstYearEdit = strEdit
End Function

'-----------------------------------------------------------------------------
' [元年⇒1年] for EraFormat
'-----------------------------------------------------------------------------
' (a) 編集文字には ee年 or e年 の[元号なし]のパターンも有り得るので
'     ターゲットを[元号+"元年"]等にしては駄目
' (b) ee年⇒"01年" , e年⇒"1年"
'
' (補) [和暦年]編集文字は e / ee のみ。eee は無い。
'      eee年は [ee]+[e年]と解釈(1年⇒ 011年 になる)される。
'-----------------------------------------------------------------------------
Private Function prvFormat_GannenTo1st(ByVal EditDate As String, _
                                       ByVal EditPattern As String) As String
Dim strEdit As String

    If (LCase(EditPattern) Like "*ee年*") Then
        strEdit = Replace(EditDate, "元年", "01年")
    Else
        strEdit = Replace(EditDate, "元年", "1年")
    End If
    
    prvFormat_GannenTo1st = strEdit
End Function

'-----------------------------------------------------------------------------
' [01年 or 1年⇒元年] for EraFormat
'-----------------------------------------------------------------------------
' (a) 編集文字には ee年 or e年 の[元号なし]のパターンも有り得るので
'     ターゲットを[元号+1年]等にしては駄目
' (b) "yyyy年(gggee年)"のように[西暦年]と併記のパターンでも
'     この関数が呼び出されるのは元年(1868,1912,1926,1989,2019)の場合だけなので
'     [西暦年]部分が "01年"/"1年"になることはない。
'
' (補) [和暦年]編集文字は e / ee のみ。eee は無いので、"001年"の変換は不要。
'      eee年は [ee]+[e年]と解釈(1年⇒ 011年 になる)される。
'-----------------------------------------------------------------------------
Private Function prvFormat_1stToGannen(ByVal EditDate As String) As String
Dim strEdit As String

    ' "01年"の置換 ⇒ "1年"の置換 の順で行なう事
    strEdit = Replace(EditDate, "01年", "元年")
    strEdit = Replace(strEdit, "1年", "元年")
    
    prvFormat_1stToGannen = strEdit
End Function


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/    [新元号:令和]対応 日付変換関数 EraCDate    ( 日付文字列 ⇒ シリアル値 )
'_/
'_/    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'_/    新元号に対応していないシステム(Office2007以前 or 対応
'_/    アップデートを施していないOffice2010以降)でも、新元号に
'_/    基づく和暦日付を日付データ(シリアル値)に変換できる関数です。
'_/
'_/    ExcelのDATEVALUE関数/VBAのCDate/DateValue関数の代わりに使用してください。
'_/
'_/    新元号に対応済みのシステムで使用しても問題ありません。
'_/
'_/    尚、EraCDate は AddinBox/kt関数アドイン(Ver5.30)に
'_/    ktEraCDate の名前で収録します。
'_/    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'_/   【 EraCDate でサポートする日付文字列のフォーマット 】
'_/      a) 区切り形式(和暦年は1～3桁,4桁以上はエラー)
'_/           [H31年4月30日] [H31/4/30] [H31-4-30] [H31.4.30]
'_/      b) 元号
'_/           [明治,明,M/m] [大正,大,T/t] [昭和,昭,S/s] [平成,平,H/h] [令和,令,R/r]
'_/      c) 西暦年(4桁,3桁,2桁)も変換可能です(2桁は2000年代と解釈します)
'_/           [2019年4月30日] [2019/4/30] [2019-4-30] [2019.4.30]
'_/         尚、西洋式の[月/日/年 or 日/月/年]フォーマットはNGです
'_/      d) 平成32 等の改元以後の年数でもＯＫとしています
'_/         但し、[元号範囲=True]指定の場合は元号期間内の日付のみが OK となります。
'_/      e) "明治元年","大正元年","昭和元年","平成元年","令和元年"という表記も可とします。
'_/      f) 時刻文字列が続いている場合、時刻込みで変換します。
'_/         但し、その時刻文字列がVBAのCDate関数で変換可能なフォーマットに限ります。
'_/
'_/    (注) 日付文字列に[1900/1/1～1900/2/29]の期間の日付を指定した場合に得られる
'_/         シリアル値は、シート上に入力した場合の値とは１日ズレます。
'_/      EraCDate : 1900/1/1⇒2, 1900/2/28⇒60, 1900/2/29⇒#VALUE!, 1900/3/1⇒61
'_/      シート上 : 1900/1/1⇒1, 1900/2/28⇒59, 1900/2/29⇒60     , 1900/3/1⇒61
'_/      (Excelは Lotus1-2-3互換の為に[1900/1/1～1900/2/29]のシリアル値を敢えてズラしています)
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Function EraCDate(ByVal 日付文字列 As String, _
                         Optional ByVal 元号範囲 As Boolean = False) As Variant

'※ 引数[日付文字列]にDate型データを指定した場合、
'   日付文字列変換されて【"yyyy/m/d"形式の日付文字列】として受け取る。
'   Date型データ：[Date型の変数][DateValue/DateSerial/CDateの結果][日付書式のセル値]
'   [日付書式のセル値]は和暦書式であっても "yyyy/m/d"形式となって受け取る
'
'※ 引数[日付文字列]に数値 or 数字文字列を指定した場合、
'   Date型の範囲 [100/1/1(-657434)～9999/12/31(2958465)] 内であればシリアル値として返す。
'
'※ 引数[日付文字列]に「空セル」を指定した場合は「空文字」として受け取るのでエラーになる。

Dim strDate As String
Dim dtmDate As Date
Dim strTemp As String
Dim intEra As Integer   ' 0:西暦, 1:明治, 2:大正, 3:昭和, 4:平成, 5:令和
Dim aryEraRange As Variant

Const cstPattern As String = "[αβγδε明大昭平令MTSHRmtshr]"

    ' "平成元年","平元年","H元年" 等の"元年"を"1年"表記に改める
    ' 元号が付かない"元年"のみは年代が特定できない為、変換エラーです
    strTemp = prvCDate_GannenTo1st(日付文字列, "明治", "明", "M")
    strTemp = prvCDate_GannenTo1st(strTemp, "大正", "大", "T")
    strTemp = prvCDate_GannenTo1st(strTemp, "昭和", "昭", "S")
    strTemp = prvCDate_GannenTo1st(strTemp, "平成", "平", "H")
    strTemp = prvCDate_GannenTo1st(strTemp, "令和", "令", "R")

    ' 元号(２文字)の判定処理が Like 演算子で行なえる様に代替文字(１文字)で置換する
    strTemp = Replace(strTemp, "明治", "α")
    strTemp = Replace(strTemp, "大正", "β")
    strTemp = Replace(strTemp, "昭和", "γ")
    strTemp = Replace(strTemp, "平成", "δ")
    strTemp = Replace(strTemp, "令和", "ε")

    '=== 日付文字列のパターンチェック ===
    '=== 年はここで数字判定を完了(# ～ #### パターン)する
    '=== 月日の数字判定はCDateに任せる(最低限,1桁数字は保証)
    '=== [日付文字列＋時刻文字列]も対象とする(末尾の*で時刻文字列 等が有ってもＯＫになる)

    '--- 和暦(年1桁) ---
    If (strTemp Like cstPattern & "#年#*月#*日*") Or _
       (strTemp Like cstPattern & "#/#*/#*") Or _
       (strTemp Like cstPattern & "#.#*.#*") Or _
       (strTemp Like cstPattern & "#-#*-#*") Then
        If (Mid(strTemp, 2, 1) = "0") Then
            EraCDate = CVErr(xlErrValue)  ' 0年 はエラー
            Exit Function
        Else
            '[元号+和暦年]⇒[西暦年]変換
            strDate = prvEraYear4EraCDate(Mid(strTemp, 1, 1), Mid(strTemp, 2, 1), Mid(strTemp, 3), intEra)
        End If

    '--- 和暦(年2桁) ---
    ElseIf (strTemp Like cstPattern & "##年#*月#*日*") Or _
           (strTemp Like cstPattern & "##/#*/#*") Or _
           (strTemp Like cstPattern & "##.#*.#*") Or _
           (strTemp Like cstPattern & "##-#*-#*") Then
        If (Mid(strTemp, 2, 2) = "00") Then
            EraCDate = CVErr(xlErrValue)    ' 00年 はエラー
            Exit Function
        Else
            '[元号+和暦年]⇒[西暦年]変換
            strDate = prvEraYear4EraCDate(Mid(strTemp, 1, 1), Mid(strTemp, 2, 2), Mid(strTemp, 4), intEra)
        End If

    ' (補) [和暦年]編集文字は e / ee のみ。eee は無いので、和暦(年3桁)の判定は不要。
    '      eee年は [ee]+[e年]と解釈(1年⇒ 011年 になる)される。

    '--- 西暦(年4桁) ---
    ' 西洋式の[月/日/年 or 日/月/年]フォーマットはNGです
    ElseIf (日付文字列 Like "####年#*月#*日*") Or _
           (日付文字列 Like "####/#*/#*") Or _
           (日付文字列 Like "####.#*.#*") Or _
           (日付文字列 Like "####-#*-#*") Then
        'ピリオド区切りがDateValueでは変換できないので / に置換する
        strDate = Replace(日付文字列, ".", "/")
        intEra = 0

    '--- 西暦(年3桁) --- (そのまま 100～999年と解釈する)
    ElseIf (日付文字列 Like "###年#*月#*日*") Or _
           (日付文字列 Like "###/#*/#*") Or _
           (日付文字列 Like "###.#*.#*") Or _
           (日付文字列 Like "###-#*-#*") Then
        'ピリオド区切りがDateValueでは変換できないので / に置換する
        strDate = Replace(日付文字列, ".", "/")
        intEra = 0

    '--- 西暦(年2桁) --- (2000年代と解釈する)
    ElseIf (日付文字列 Like "##年#*月#*日*") Or _
           (日付文字列 Like "##/#*/#*") Or _
           (日付文字列 Like "##.#*.#*") Or _
           (日付文字列 Like "##-#*-#*") Then
        'ピリオド区切りがDateValueでは変換できないので / に置換する
        strDate = "20" & Replace(日付文字列, ".", "/")  '先頭に"20"を付加して2000年代
        intEra = 0

    '--- 数値 or 数字文字列 ---
    ' ここで、即 CDate変換(単なるDate型への型変換)して値を返す
    ElseIf IsNumeric(日付文字列) Then
        '※ カンマ編集数値の場合はカンマを取り除く
        '   カンマも日付区切り文字として扱われてしまい、
        '   省略形の日付文字列として予想外の日付と見做される
        '   CDate("2,500")⇒500/2/1 と解釈(m,yyy)される
        On Error Resume Next
        dtmDate = CDate(Replace(日付文字列, ",", ""))
        If (Err.Number <> 0) Then
            'シリアル値範囲外 [100/1/1(-657434)～9999/12/31(2958465)]
            EraCDate = CVErr(xlErrValue)
        Else
            EraCDate = dtmDate
        End If
        On Error GoTo 0
        Exit Function

    '--- 他はエラー ---
    Else
        EraCDate = CVErr(xlErrValue)
        Exit Function
    End If

    '=== CDateによる日付文字列(時刻文字列を含んでも可)⇒シリアル値 変換 ===
    ' 和暦年は西暦年に変換済
    On Error Resume Next
    dtmDate = CDate(strDate)
    If (Err.Number <> 0) Then
        EraCDate = CVErr(xlErrValue)
        Exit Function
    End If
    On Error GoTo 0

    If (元号範囲 = True) Then
        If (intEra = 0) Then
            '西暦は範囲無し
        Else
            ' 明治[1]：1868(M1)/10/23 ～ 1912(M45)/ 7/29
            ' 大正[2]：1912(T1)/ 7/30 ～ 1926(T15)/12/24
            ' 昭和[3]：1926(S1)/12/25 ～ 1989(S64)/ 1/ 7
            ' 平成[4]：1989(H1)/ 1/ 8 ～ 2019(H31)/ 4/30
            ' 令和[5]：2019(R1)/ 5/ 1 ～ 9999(---)/12/31
            ' ※[時刻]も指定されている場合には小数があるので、
            '   終端の判定は[終了日＋１より小]とする
            aryEraRange = _
                Array(Array(0, 0), Array(#10/23/1868#, #7/29/1912#), _
                      Array(#7/30/1912#, #12/24/1926#), Array(#12/25/1926#, #1/7/1989#), _
                      Array(#1/8/1989#, #4/30/2019#), Array(#5/1/2019#, #12/31/9999#))
            If (aryEraRange(intEra)(0) <= dtmDate) And _
               ((aryEraRange(intEra)(1) + 1) > dtmDate) Then
                '元号範囲内でＯＫ
            Else
                EraCDate = CVErr(xlErrValue)
                Exit Function
            End If
        End If
    End If

    EraCDate = dtmDate

End Function

'-----------------------------------------------------------------------------
' [元年⇒1年] for EraCDate   Era1:令和, Era2:令, Era3:R 等
'-----------------------------------------------------------------------------
Private Function prvCDate_GannenTo1st _
            (ByVal EditDate As String, ByVal Era1 As String, _
             ByVal Era2 As String, ByVal Era3 As String) As String
Dim strEdit As String

    strEdit = Replace(EditDate, (Era1 & "元年"), (Era1 & "1年"))
    strEdit = Replace(strEdit, (Era2 & "元年"), (Era2 & "1年"))
    strEdit = Replace(strEdit, (Era3 & "元年"), (Era3 & "1年"))
    strEdit = Replace(strEdit, (LCase(Era3) & "元年"), (Era3 & "1年"))
    
    prvCDate_GannenTo1st = strEdit
End Function

'-----------------------------------------------------------------------------
' [元号+和暦年] ⇒ [西暦年]変換 , 元号を[元号フラグ]で返す
'
' ２文字元号はα(明治),β(大正),γ(昭和),δ(平成),ε(令和)に置換されている
' [年]は 1～2文字の数字
' [月日]には[年]数字以降の日付文字列(時刻文字列を含む)の部分が渡される
'-----------------------------------------------------------------------------
Private Function prvEraYear4EraCDate(ByVal 元号 As String, ByVal 年 As String, _
                                     ByVal 月日 As String, ByRef 元号フラグ As Integer) As String
Dim strDate As String
Dim strMMDD As String

    'ピリオド区切りがDateValueでは変換できない為、スラッシュに置換する
    strMMDD = Replace(月日, ".", "/")

    '元号に応じて西暦年に変換する(年は Like演算により数字のチェック済)
    Select Case 元号
      Case "α", "明", "M", "m"
        strDate = (CLng(年) + 1867) & strMMDD  '明治
        元号フラグ = 1
      Case "β", "大", "T", "t"
        strDate = (CLng(年) + 1911) & strMMDD  '大正
        元号フラグ = 2
      Case "γ", "昭", "S", "s"
        strDate = (CLng(年) + 1925) & strMMDD  '昭和
        元号フラグ = 3
      Case "δ", "平", "H", "h"
        strDate = (CLng(年) + 1988) & strMMDD  '平成
        元号フラグ = 4
      Case "ε", "令", "R", "r"
        strDate = (CLng(年) + 2018) & strMMDD  '令和
        元号フラグ = 5
    End Select
    prvEraYear4EraCDate = strDate
End Function


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/    [新元号:令和]対応 日付判定関数 EraIsDate    ( 日付データ ⇒ True/False )
'_/
'_/    新元号に対応していないシステム(Office2007以前 or 対応
'_/    アップデートを施していないOffice2010以降)でも、新元号に
'_/    基づく和暦日付を含め、「日付として妥当か否か」を判定する関数です。
'_/
'_/    EraIsDate がサポートする日付文字列のフォーマットはEraCDateに準じます。
'_/
'_/    新元号に対応済みのシステムで使用しても問題ありません。
'_/
'_/    尚、EraIsDate は AddinBox/kt関数アドイン(Ver5.30)に
'_/    ktEraIsDate の名前で収録します。
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Function EraIsDate(ByVal 日付データ As Variant) As Boolean
Dim strData As String
Dim Result As Variant

    If IsError(日付データ) Then
        EraIsDate = False

    ElseIf IsEmpty(日付データ) Then  '空セルはEmpty値
        EraIsDate = False

    ElseIf IsDate(日付データ) Then
        'IsDateの判定対象：「日付文字列」「Date型のデータ」
        '(1) EraCDateでもIsDate判定が可能だが、効率の面から
        '    IsDateで対応可能な部分は先にIsDateで済ませる
        '
        '(2) [IsDate(Date型データ) ⇒ True]
        '
        '(3) [IsDate("明治33年2月29日") ⇒False]
        '    シート上でのシリアル値60に対する日付文字列
        '    EraCDateでも False になる
        '
        '(4) [IsDate("昭和65年1月1日")⇒False (元号範囲オーバー)] … 後のEraCDateで救済
        '    [IsDate(数値)⇒False] … 次の IsNumericで救済
        '
        '(5) VBAなのでマイナスのシリアル値に対応する
        '    日付文字列("M1/10/23" , "1899/1/1" 等)も True になる
        '
        '(6) [令和]日付文字列
        '      新元号アップデート済 ⇒ True
        '      新元号アップデート未 ⇒ False … 後のEraCDateで救済
        '
        '(7) "平成元年1月8日"等の[元年表記]
        '      新元号アップデート済 ⇒ True
        '      新元号アップデート未 ⇒ False … 後のEraCDateで救済

        EraIsDate = True

    ElseIf IsNumeric(日付データ) Then
        'IsNumericの判定対象：「数値 , 数字文字列」
        '(1) 単なる数値判定なのでマイナス値(1899年以前の日付)も True になる
        '
        '(2) シリアル値の範囲は 100/1/1(-657434)～9999/12/31(2958465)
        
        If (VarType(日付データ) = vbString) Then
            If (CDbl(日付データ) >= -657434) And (CDbl(日付データ) <= 2958465) Then
                EraIsDate = True
            Else
                EraIsDate = False
            End If
        ElseIf (日付データ >= -657434) And (日付データ <= 2958465) Then
            EraIsDate = True
        Else
            EraIsDate = False
        End If

    Else
        'Elseの判定対象：
        '  数値/文字列以外のデータ型⇒False
        '  IsDate で弾かれた日付文字列(下記)の救済⇒True
        '(1) [令和]日付文字列
        '(2) "平成元年1月8日"等の[元年表記]
        '(3) "昭和65年1月1日"等の元号範囲オーバー
        On Error Resume Next
        strData = CStr(日付データ)
        If (Err.Number <> 0) Then
            EraIsDate = False   '数値/文字列以外のデータは対象外
            Exit Function
        End If
        On Error GoTo 0
        
        Result = EraCDate(strData, False)   'EraCDateで検証する(元号範囲は無視)
        If IsError(Result) Then
            EraIsDate = False
        Else
            EraIsDate = True
        End If
    End If
End Function

