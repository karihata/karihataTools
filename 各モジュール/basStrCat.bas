Attribute VB_Name = "basStrCat"
Option Explicit

'�z��𕶎���������֐�
Function StrCat(Elements As Variant, Optional separator As String = vbNullString) As String
Attribute StrCat.VB_Description = "�͈͑I���������̂��u�����N�Z�����Ȃ��ĕ������������܂��B"
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
