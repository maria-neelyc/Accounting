Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private dbs As Database
Private rst As String
Private fld As String
Sub Class_Initialize()
    Set dbs = CurrentDb
End Sub
Sub Class_Terminate()
    Set dbs = Nothing
End Sub
Sub Location(RecordsetName As String, fieldName As String)
    rst = RecordsetName
    fld = fieldName
End Sub
Property Let Value(lngVal As Long)
On Error GoTo Err_End:
    Dim tmprst As Recordset
    Dim tmpVal As Long
    'Set tmprst = dbs.OpenRecordset(rst, dbOpenDynaset, dbDenyWrite)
    Set tmprst = dbs.OpenRecordset(rst)
    tmprst.MoveFirst
    tmprst.Edit
    tmprst(fld).Value = lngVal
    tmprst.Update
Err_Exit:
    Exit Property
Err_End:
    MsgBox "Set Location First!", vbCritical
    GoTo Err_Exit
End Property
Property Get Value() As Long
On Error GoTo Err_End:
    Dim tmprst As Recordset
    Dim tmpVal As Long
    Set tmprst = dbs.OpenRecordset(rst, dbOpenSnapshot)
    tmprst.MoveFirst
    Value = tmprst(fld).Value
Err_Exit:
    Exit Property
Err_End:
    MsgBox "NoData in specific Field!", vbCritical
    GoTo Err_Exit
End Property