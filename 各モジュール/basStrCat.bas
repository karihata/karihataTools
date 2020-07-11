Attribute VB_Name = "basStrCat"
Option Explicit

'配列を文字結合する関数
Function StrCat(Elements As Variant, Optional separator As String = vbNullString) As String
Attribute StrCat.VB_Description = "範囲選択したものをブランクセルを省いて文字を結合します。"
Attribute StrCat.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim output As String, tmp As String, e As Variant
    output = vbNullString

    For Each e In Elements
        tmp = CStr(e)
        If 0 < Len(tmp) Then
            output = output & tmp & separator
        End If
    Next

    StrCat = Left(output, Len(output) - Len(separator))
End Function
