Attribute VB_Name = "basPush"
Option Explicit

' VariantŒ^‚Ì”z—ñ‚É—v‘f‚ð’Ç‰Á‚·‚é
Sub Push(ByRef v As Variant, e)
    ReDim Preserve v(UBound(v) + 1)
    v(UBound(v)) = e
End Sub
