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
 
  MsgBox "変換するエクセルファイルがあるフォルダを選択してください。" & vbCrLf & "(.xlsxのみが対象です)"
  '変換元のエクセルがあるフォルダをダイアログで指定してもらう
  With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = True Then
        Search_Folder = .SelectedItems(1)
    Else
        'ダイアログでキャンセルもしくは閉じられた場合は処理を中止
        Exit Sub
    End If
  End With
  
  MsgBox "変換後のPDFファイルの保存先フォルダを選択してください。"
  '変換後のPDFファイルを保存するフォルダをダイアログで指定してもらう
  With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = True Then
        Save_Folder = .SelectedItems(1)
    Else
        'ダイアログでキャンセルもしくは閉じられた場合は処理を中止
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
    MsgBox (".xlsxファイルが見つからないため終了します。")
  Else
    rc = MsgBox(".xlsxファイルが" & cnt & "件見つかりました。一括変換処理を行いますか？", vbYesNo + vbQuestion, "確認")
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
      MsgBox ("処理が完了しました。")
    Else
      MsgBox ("処理を中断しました。")
    End If
  End If

  Set objFs = Nothing
End Sub


