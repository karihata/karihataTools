Attribute VB_Name = "basStrCatIf"
Option Explicit

'ğŒˆê’v‚µ‚½”z—ñ‚ğ•¶šŒ‹‡‚·‚éŠÖ”
Function StrCatIf(TermElements As Variant, Terms As String, Elements As Variant, Optional separator As String = vbNullString) As String
    Dim output As String, tmp As String, e As Variant, i As Long
    Dim outputElements
    
    'Push‚·‚é‚½‚ß‚É”z—ñ‚Å‰Šú‰»
    outputElements = Array()
    output = vbNullString
    
    'ğŒˆê’v‚µ‚Ä‚¢‚é‚à‚Ì‚ğ•Ê‚Ì”z—ñ‚ÉˆÚ‚·
    For i = 1 To TermElements.Count
        If Terms = TermElements(i) Then
            Push outputElements, Elements(i)
        End If
    Next i
   
    
    '•¶šŒ‹‡‚ğ‚·‚é
    '•¶š”0‚Í–³‹‚·‚é
    For Each e In outputElements
        tmp = CStr(e)
        If 0 < Len(tmp) Then
            output = output & tmp & separator
        End If
    Next

    StrCatIf = Left(output, Len(output) - Len(separator))
End Function

