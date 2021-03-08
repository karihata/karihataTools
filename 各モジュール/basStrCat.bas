Attribute VB_Name = "basStrCat"
Option Explicit

'配列を文字結合する関数
'Elements As Variant　　　　　　　　　　　　　　文字結合したいセル範囲
'Optional separator As String = vbNullString　区切り文字、引数に入れない場合は区切り文字は入らない
'セル範囲でブランクセルは文字結合から省かれて区切り文字も入らない
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
