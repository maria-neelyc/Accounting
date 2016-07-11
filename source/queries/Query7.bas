Operation =1
Option =0
Begin InputTables
    Name ="tblAccount"
    Alias ="cAcc"
    Name ="prevtblAccount"
    Alias ="pACC"
End
Begin OutputColumns
    Expression ="cAcc.accID"
    Expression ="cAcc.accName"
    Expression ="cAcc.accNo"
    Expression ="pACC.accNo"
    Alias ="ssign"
    Expression ="IIf([cacc].[accsign]<>[pacc].[accsign],\"Diff\")"
    Alias ="etype"
    Expression ="IIf([cacc].[acctype]<>[pacc].[acctype],\"Diff\")"
    Expression ="cAcc.accOpenBalCur"
    Expression ="cAcc.accOpenBalance"
    Expression ="pACC.accOpenBalance"
    Expression ="cAcc.accType"
    Expression ="pACC.accType"
End
Begin Joins
    LeftTable ="cAcc"
    RightTable ="pACC"
    Expression ="cAcc.accID = pACC.accID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
Begin
    State =0
    Left =18
    Top =8
    Right =1898
    Bottom =752
    Left =-1
    Top =-1
    Right =1859
    Bottom =411
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =7
        Right =196
        Bottom =333
        Top =0
        Name ="cAcc"
        Name =""
    End
    Begin
        Left =464
        Top =11
        Right =661
        Bottom =331
        Top =0
        Name ="pACC"
        Name =""
    End
End
