Operation =1
Option =0
Begin InputTables
    Name ="tblControl"
    Name ="sysMyCo"
End
Begin OutputColumns
    Expression ="sysMyCo.smycoPath"
End
Begin Joins
    LeftTable ="tblControl"
    RightTable ="sysMyCo"
    Expression ="tblControl.ctrDBID = sysMyCo.smycoID"
    Flag =1
End
Begin Groups
    Expression ="sysMyCo.smycoPath"
    GroupLevel =0
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
    State =1
    Left =0
    Top =783
    Right =160
    Bottom =817
    Left =-1
    Top =-1
    Right =1841
    Bottom =315
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =281
        Top =0
        Name ="tblControl"
        Name =""
    End
    Begin
        Left =216
        Top =7
        Right =336
        Bottom =155
        Top =0
        Name ="sysMyCo"
        Name =""
    End
End
