Option Compare Database
Public bolLOG
Function FnLog(strVar As Variant)
    If bolLOG = True Then
        Open Left(CurrentDb.name, InStrRev(CurrentDb.name, "\")) & "EasyBookQuery.txt" For Append As #1 'was For Output
        Write #1, Application.CurrentObjectName
        'Write #1,
        Write #1, strVar
        Write #1,
        Close #1
    End If
End Function
Function SetLog()
    If bolLOG = True Then
        bolLOG = False
        Application.CommandBars("Ledger  EasyBook").Controls("Tools").Controls("QueryTrace Off").Caption = "QueryTrace On"
    Else
        bolLOG = True
        Application.CommandBars("Ledger  EasyBook").Controls("Tools").Controls("QueryTrace On").Caption = "QueryTrace Off"
    End If
    
'    Select Case bolLOG
'        Case True: bolLOG = False
'                   Application.CommandBars("Ledger  EasyBook").Controls("Tools").Controls("QueryTrace Off").Caption = "QueryTrace On"
'        Case False: bolLOG = True
'                   Application.CommandBars("Ledger  EasyBook").Controls("Tools").Controls("QueryTrace On").Caption = "QueryTrace Off"
'    End Select
End Function

Function FixToolsMenu()
    If IsEmpty(bolLOG) Then
        On Error Resume Next
        Application.CommandBars("Ledger  EasyBook").Controls("Tools").Controls("QueryTrace Off").Caption = "QueryTrace On"
        On Error Resume Next
        Application.CommandBars("Ledger  EasyBook").Controls("Tools").Controls("QueryTrace On").Caption = "QueryTrace On"
        bolLOG = False
    End If
End Function