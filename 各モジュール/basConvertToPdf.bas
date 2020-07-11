Attribute VB_Name = "basConvertToPdf"
Option Explicit

Sub ConvertToPdf(ByVal control As IRibbonControl)
  Dim buf As String
  Dim cnt As Long
  Dim files() As String
  Dim rc As Integer
  Dim item As Variant
  Dim fullname As String
  Dim fullnamePdf As String
  Dim objExcel As Object 'Excel.Application
  Dim objBook As Object 'Excel.Workbook
  Dim objFs As New Scripting.FileSystemObject 'Scripting.FilesystemObject
  Dim Search_Folder As String
  Dim Save_Folder As String
 
  MsgBox "�ϊ�����G�N�Z���t�@�C��������t�H���_��I�����Ă��������B" & vbCrLf & "(.xlsx�݂̂��Ώۂł�)"
  '�ϊ����̃G�N�Z��������t�H���_���_�C�A���O�Ŏw�肵�Ă��炤
  With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = True Then
        Search_Folder = .SelectedItems(1)
    Else
        '�_�C�A���O�ŃL�����Z���������͕���ꂽ�ꍇ�͏����𒆎~
        Exit Sub
    End If
  End With
  
  MsgBox "�ϊ����PDF�t�@�C���̕ۑ���t�H���_��I�����Ă��������B"
  '�ϊ����PDF�t�@�C����ۑ�����t�H���_���_�C�A���O�Ŏw�肵�Ă��炤
  With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = True Then
        Save_Folder = .SelectedItems(1)
    Else
        '�_�C�A���O�ŃL�����Z���������͕���ꂽ�ꍇ�͏����𒆎~
        Exit Sub
    End If
  End With
  
  On Error Resume Next
  
  buf = Dir(Search_Folder & "\*.xlsx")
  cnt = 0
  Do While buf <> ""
    ReDim Preserve files(cnt)
    files(cnt) = buf
    cnt = cnt + 1
    buf = Dir()
  Loop

  If cnt = 0 Then
    MsgBox (".xlsx�t�@�C����������Ȃ����ߏI�����܂��B")
  Else
    rc = MsgBox(".xlsx�t�@�C����" & cnt & "��������܂����B�ꊇ�ϊ��������s���܂����H", vbYesNo + vbQuestion, "�m�F")
    If rc = vbYes Then
      Set objExcel = CreateObject("Excel.Application")
      objExcel.Visible = False
      For Each item In files()
        fullname = Search_Folder & "\" & item
        fullnamePdf = Save_Folder & "\" & objFs.GetBaseName(item) & ".pdf"
        Set objBook = objExcel.Workbooks.Open(fullname, , True)
        objBook.ExportAsFixedFormat 0, fullnamePdf
        objBook.Close (False)
        Set objBook = Nothing
      Next item
      objExcel.Quit
      Set objExcel = Nothing
      MsgBox ("�������������܂����B")
    Else
      MsgBox ("�����𒆒f���܂����B")
    End If
  End If

  Set objFs = Nothing
End Sub


