Operation =1
Option =0
Where ="(((tblTransactionSub.accID)=129))"
Begin InputTables
    Name ="tblTransaction"
    Name ="tblTransactionSub"
End
Begin OutputColumns
    Expression ="tblTransactionSub.*"
End
Begin Joins
    LeftTable ="tblTransaction"
    RightTable ="tblTransactionSub"
    Expression ="tblTransaction.trnID = tblTransactionSub.trnID"
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
    Bottom =381
    Left =-1
    Top =-1
    Right =1859
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =7
        Right =238
        Bottom =155
        Top =0
        Name ="tblTransaction"
        Name =""
    End
    Begin
        Left =398
        Top =3
        Right =599
        Bottom =151
        Top =0
        Name ="tblTransactionSub"
        Name =""
    End
End
