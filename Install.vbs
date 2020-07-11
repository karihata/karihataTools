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

    MsgBox("Microsoft�I�t�B�X��bit�����m�F���ĉ������B" + vbCrLf + "ImageMagick���C���X�g�[������ۂ�bit���ɂ��������̂ɂ��ĉ������B")

    ' �C���g�[���m�F�E�B���h�E
    IF MsgBox("�A�h�C�����C���g�[�����܂����H", vbYesNo + vbQuestion) = vbNo Then
        WScript.Quit
    End IF

    ' Excel�C���^���X��
    Set objExcel   = CreateObject("Excel.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    ' �p�X�ݒ�
    strAdPath = objExcel.Application.UserLibraryPath
    strMyPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    strAdCp   = objFileSys.BuildPath(strAdPath, FILLE_NAME)
    strMyCp   = objFileSys.BuildPath(strMyPath, FILLE_NAME)

    ' �t�@�C���R�s�[
    objFileSys.CopyFile strMyCp, strAdCp

    ' �A�h�C���o�^
    objExcel.Workbooks.Add
    Set oAdd = objExcel.AddIns.Add(strAdCp,True)
    ' �A�h�C���L����
    oAdd.Installed = True
    objExcel.Quit

    Set objExcel   = Nothing
    Set objFileSys = Nothing

    MsgBox "�C���g�[�����������܂���"
End Sub