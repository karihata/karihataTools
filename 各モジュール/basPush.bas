Attribute VB_Name = "basPush"
Option Explicit

' Variant型の配列に要素を追加する
Sub Push(ByRef v As Variant, e)
    ReDim Preserve v(UBound(v) + 1)
    v(UBound(v)) = e
End Sub
