Attribute VB_Name = "basPush"
Option Explicit

' Variant�^�̔z��ɗv�f��ǉ�����
Sub Push(ByRef v As Variant, e)
    ReDim Preserve v(UBound(v) + 1)
    v(UBound(v)) = e
End Sub
