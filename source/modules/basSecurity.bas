Option Compare Database

Function DiskZeroName() As String
    Dim buffer As String
    Dim PathTxt As String
    Dim FileNumber, i As Integer
    On Error Resume Next
    Kill "DiskZeroName*.*"
    On Error GoTo 0
    PathTxt = "DiskZeroName" & Format(Now(), "dddmmmddyyyyhhnnss\.\t\x\t")
    Shell "RegEdit /E " & PathTxt & " HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\Scsi\", vbHide
    FileNumber = FreeFile
    Open PathTxt For Binary As #FileNumber
    buffer = String(LOF(FileNumber), vbNullChar)
    Get #FileNumber, , buffer
    Close #FileNumber
    buffer = StrConv(buffer, vbFromUnicode)
    DiskZeroName = Mid$(buffer, InStr(buffer, "Identifier") + 13)
    DiskZeroName = Left$(DiskZeroName, InStr(DiskZeroName, Chr(34)) - 1)
    Kill "DiskZeroName*.*"
    buffer = ""
    For i = 1 To Len(DiskZeroName)
        If IsNumeric(Left(DiskZeroName, 1)) = True Then
            buffer = buffer & Left(DiskZeroName, 1)
        End If
        DiskZeroName = Right(DiskZeroName, Len(DiskZeroName) - 1)
    Next i
    DiskZeroName = Hex(buffer)
    If DiskZeroName <> "203FEF24" Then '203FEF24
        MsgBox "This application was not properly installed. Please contact NeeLyc.", vbCritical, "Security Error"
        'Application.CloseCurrentDatabase
    End If
End Function

Public Function GetHardDriveSerialNumber()
     MsgBox Hex(CreateObject("Scripting.FileSystemObject").GetDrive(Left(Application.CurrentProject.Path, InStr(Application.CurrentProject.Path, "\"))).SerialNumber)
End Function
Function AddAppProperty(strName As String, varType As Variant, varValue As Variant) As Integer
    Dim dbs As Database, prp As Property
    Const conPropNotFoundError = 3270

    Set dbs = CurrentDb
    On Error GoTo AddProp_Err
    dbs.Properties(strName) = varValue

AddAppProperty = True

AddProp_Bye:
    Exit Function

AddProp_Err:
    If Err = conPropNotFoundError Then
        Set prp = dbs.CreateProperty(strName, varType, varValue)
        dbs.Properties.Append prp
        Resume
    Else
        AddAppProperty = False
        Resume AddProp_Bye
    End If
End Function