VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'使い方-------------------------------------------------------
'dim objpb

' 進捗バークラスの定義
'Set objpb = New ProgressBar

' 進捗バーのタイトルを設定
'objpb.SetTitle "パス確認"

'プログレスバー更新
'objpb.SetTitle "パス確認　" & i & "：" & Selection.Count
'objpb.SetProgress i / Selection.Count

'プログレスバーの破棄
'Set objpb = Nothing
'-------------------------------------------------------------

'-------------------------------------------------------------------------------
' ProgressBar : 進捗バークラス
'-------------------------------------------------------------------------------
Const WIDTH = 400
Const HEIGHT = 150
Const BAR_WIDTH = 350
Const BAR_HEIGHT = 16
Const BAR_BG = "#C0C0C0"
Const BAR_FG = "#0066FF"


'----------------------------------------------------
' クラス変数定義
'----------------------------------------------------
Private strTitle1
Private nCurrent1
Private nStartTime, nCurrentTime
Private objIE
Private div1, div2, div3

'----------------------------------------------------
' コンストラクタ
'----------------------------------------------------
Private Sub Class_Initialize()
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = False
    objIE.Navigate2 "about:blank"
    objIE.document.Title = "進捗状況"
    objIE.AddressBar = False
    objIE.MenuBar = False
    objIE.Toolbar = False
    objIE.Resizable = False
    objIE.WIDTH = WIDTH
    objIE.HEIGHT = HEIGHT
    objIE.Top = 0
    objIE.Left = 0
    Set div1 = objIE.document.createElement("div")
    div1.ID = "div1"
    div1.Style.Position = "absolute"
    div1.Style.Top = "10px"
    div1.Style.Left = "10px"
    div1.Style.backgroundColor = BAR_BG
    div1.Style.WIDTH = BAR_WIDTH & "px"
    div1.Style.HEIGHT = BAR_HEIGHT & "px"
    div1.Style.Border = "1px solid"
    div1.Style.overflow = "hidden"
    Set div2 = objIE.document.createElement("div")
    div2.ID = "div2"
    div2.Style.Position = "relative"
    div2.Style.Top = "1px"
    div2.Style.Left = "1px"
    div2.Style.backgroundColor = BAR_FG
    div2.Style.WIDTH = "0px"
    div2.Style.HEIGHT = (BAR_HEIGHT - 2) & "px"
    div2.Style.overflow = "hidden"
    Set div3 = objIE.document.createElement("div")
    div3.ID = "div3"
    div3.Style.Position = "absolute"
    div3.Style.Top = "45px"
    div3.Style.Left = "10px"

    objIE.document.body.appendChild (div1)
    div1.appendChild (div2)
    objIE.document.body.appendChild (div3)

    nStartTime = Timer()
End Sub

'----------------------------------------------------
' デストラクタ
'----------------------------------------------------
Private Sub Class_Terminate()
    On Error Resume Next
    objIE.Quit
    Set objIE = Nothing
    On Error GoTo 0
End Sub

'----------------------------------------------------
' 表示タイトルの設定
'----------------------------------------------------
Public Sub SetTitle(t)
    strTitle1 = t
    objIE.document.Title = t & String(40, "　")
    objIE.Visible = True
End Sub

'----------------------------------------------------
' 進捗パーセントの設定
'----------------------------------------------------
Public Sub SetProgress(n1)
    nCurrent1 = n1
    Repaint
End Sub

'----------------------------------------------------
' 進捗バー再描画
'----------------------------------------------------
Private Sub Repaint()

Dim nAverage
Dim nElapsedTime
Dim nRemain
Dim strRemain
Dim strPercent

Dim w1
Dim style1, style2

    nCurrentTime = Timer()
    nElapsedTime = nCurrentTime - nStartTime

    w1 = BAR_WIDTH * (nCurrent1)
    strRemain = "不明"
    If nElapsedTime <> 0 Then
        nAverage = nCurrent1 / nElapsedTime
        If nAverage <> 0 Then
            nRemain = Round((1 - nCurrent1) / nAverage, 1)
        End If
        If nRemain > 60 Then
            strRemain = "約" & CStr(Round(nRemain / 60, 0)) & "分"
        Else
            strRemain = FormatNumber(nRemain, 1) & "秒"
        End If
    End If

    strPercent = FormatNumber(nCurrent1 * 100, 1)

    On Error Resume Next
    div2.Style.WIDTH = (w1 - 1) & "px"
    div3.innerText = strPercent & "%終了　--　残り推定：" & strRemain
    objIE.Visible = True
    objIE.document.all(0).Click
    On Error GoTo 0
End Sub
