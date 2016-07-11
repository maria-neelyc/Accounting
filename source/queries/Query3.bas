Operation =1
Option =0
Begin InputTables
    Name ="tblControl"
    Name ="sysMyCo"
End
Begin OutputColumns
    Expression ="tblControl.*"
    Expression ="sysMyCo.smycoName"
    Expression ="sysMyCo.smycoPath"
    Expression ="sysMyCo.smycoYear"
End
Begin Joins
    LeftTable ="tblControl"
    RightTable ="sysMyCo"
    Expression ="tblControl.ctrDBID = sysMyCo.smycoID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
Begin
End
Begin
    State =0
    Left =18
    Top =41
    Right =1221
    Bottom =353
    Left =-1
    Top =-1
    Right =1196
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =118
        Top =7
        Name ="tblControl"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="sysMyCo"
        Name =""
    End
End
