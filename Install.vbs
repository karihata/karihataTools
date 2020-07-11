Const FILLE_NAME="KarihataTools.xlam"

Call Exec

Sub Exec()
    Dim objExcel
    Dim strAdPath
    Dim strMyPath
    Dim strAdCp
    Dim strMyCp
    Dim objFileSys
    Dim oAdd

    MsgBox("Microsoftオフィスのbit数を確認して下さい。" + vbCrLf + "ImageMagickをインストールする際はbit数にあったものにして下さい。")

    ' イントール確認ウィンドウ
    IF MsgBox("アドインをイントールしますか？", vbYesNo + vbQuestion) = vbNo Then
        WScript.Quit
    End IF

    ' Excelインタンス化
    Set objExcel   = CreateObject("Excel.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    ' パス設定
    strAdPath = objExcel.Application.UserLibraryPath
    strMyPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    strAdCp   = objFileSys.BuildPath(strAdPath, FILLE_NAME)
    strMyCp   = objFileSys.BuildPath(strMyPath, FILLE_NAME)

    ' ファイルコピー
    objFileSys.CopyFile strMyCp, strAdCp

    ' アドイン登録
    objExcel.Workbooks.Add
    Set oAdd = objExcel.AddIns.Add(strAdCp,True)
    ' アドイン有効化
    oAdd.Installed = True
    objExcel.Quit

    Set objExcel   = Nothing
    Set objFileSys = Nothing

    MsgBox "イントールが完了しました"
End Sub