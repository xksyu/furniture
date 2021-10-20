Attribute VB_Name = "Module2"
Sub Backup()
Attribute Backup.VB_ProcData.VB_Invoke_Func = "s\n14"
    Dim x As String
    strPath = "/Users/xksyu/Desktop/"
    On Error Resume Next
    x = GetAttr(strPath) And 0
    If Err = 0 Then
        strDate = Format(Now, "dd-mm-yy hh-mm")
        FileNameXls = strPath & "\" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1) & "Мебель_Хабибулина" & strDate & ".xls"
        ActiveWorkbook.SaveCopyAs FileName:=FileNameXls
    Else
        MsgBox "Папка " & strPath & " недоступна или не существует!", vbCritical
    End If
End Sub
