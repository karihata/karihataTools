Attribute VB_Name = "basStrCatIf"
Option Explicit

'条件一致した配列を文字結合する関数
Function StrCatIf(TermElements As Variant, Terms As String, Elements As Variant, Optional separator As String = vbNullString) As String
    Dim output As String, tmp As String, e As Variant, i As Long
    Dim outputElements
    
    'Pushするために配列で初期化
    outputElements = Array()
    output = vbNullString
    
    '条件一致しているものを別の配列に移す
    For i = 1 To TermElements.Count
        If Terms = TermElements(i) Then
            Push outputElements, Elements(i)
        End If
    Next i
   
    
    '文字結合をする
    '文字数0は無視する
    For Each e In outputElements
        tmp = CStr(e)
        If 0 < Len(tmp) Then
            output = output & tmp & separator
        End If
    Next

    StrCatIf = Left(output, Len(output) - Len(separator))
End Function

