Attribute VB_Name = "basStrCatIf"
Option Explicit

'������v�����z��𕶎���������֐�
Function StrCatIf(TermElements As Variant, Terms As String, Elements As Variant, Optional separator As String = vbNullString) As String
    Dim output As String, tmp As String, e As Variant, i As Long
    Dim outputElements
    
    'Push���邽�߂ɔz��ŏ�����
    outputElements = Array()
    output = vbNullString
    
    '������v���Ă�����̂�ʂ̔z��Ɉڂ�
    For i = 1 To TermElements.Count
        If Terms = TermElements(i) Then
            Push outputElements, Elements(i)
        End If
    Next i
   
    
    '��������������
    '������0�͖�������
    For Each e In outputElements
        tmp = CStr(e)
        If 0 < Len(tmp) Then
            output = output & tmp & separator
        End If
    Next

    StrCatIf = Left(output, Len(output) - Len(separator))
End Function

