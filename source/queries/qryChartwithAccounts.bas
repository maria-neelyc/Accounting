Operation =1
Option =0
Begin InputTables
    Name ="tblChart"
    Name ="tblAccount"
End
Begin OutputColumns
    Expression ="tblChart.coaID"
    Expression ="tblChart.lvlID"
    Expression ="tblChart.coaIDg"
    Expression ="tblChart.coaNo"
    Expression ="tblChart.coaName"
    Expression ="tblChart.coaRef1"
    Expression ="tblChart.coaRef2"
    Expression ="tblChart.coaRef3"
    Expression ="tblChart.coaRef4"
    Expression ="tblChart.coaRef5"
    Expression ="tblAccount.accNo"
    Expression ="tblAccount.accName"
    Expression ="tblAccount.accID"
    Expression ="tblAccount.accIsVat"
    Alias ="accType"
    Expression ="IIf(IsNull(tblAccount.accType),tblChart.coaAccType,tblAccount.accType)"
    Expression ="tblAccount.accOpenBalance"
    Expression ="tblAccount.accCur"
    Expression ="tblAccount.accOpenBalCur"
End
Begin Joins
    LeftTable ="tblChart"
    RightTable ="tblAccount"
    Expression ="tblChart.coaID=tblAccount.caoID"
    Flag =2
End
Begin OrderBy
    Expression ="tblChart.coaRef1"
    Flag =0
    Expression ="tblChart.coaRef2"
    Flag =0
    Expression ="tblChart.coaRef3"
    Flag =0
    Expression ="tblChart.coaRef4"
    Flag =0
    Expression ="tblChart.coaRef5"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="accType"
    End
End
Begin
    State =0
    Left =60
    Top =110
    Right =1212
    Bottom =422
    Left =-1
    Top =-1
    Right =1141
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="tblChart"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =305
        Bottom =113
        Top =8
        Name ="tblAccount"
        Name =""
    End
End
