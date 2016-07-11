Operation =1
Option =0
Where ="(((prevtblTransaction.trnYear) In (SELECT yearID FROM prevtblControl)) AND ((pre"
    "vtblAccount.accType)=2))"
Begin InputTables
    Name ="prevtblAccount"
    Name ="prevtblTransactionSub"
    Name ="prevtblTransaction"
End
Begin OutputColumns
    Alias ="Bal"
    Expression ="Sum(IIf([trnsSign]='C',-1*[trnsAmount],[trnsAmount]))+prevtblAccount.accOpenBala"
        "nce"
    Expression ="prevtblAccount.accNo"
    Expression ="prevtblAccount.caoID"
    Expression ="prevtblAccount.accName"
End
Begin Joins
    LeftTable ="prevtblAccount"
    RightTable ="prevtblTransactionSub"
    Expression ="prevtblAccount.accID = prevtblTransactionSub.accID"
    Flag =2
    LeftTable ="prevtblTransaction"
    RightTable ="prevtblTransactionSub"
    Expression ="prevtblTransaction.trnID = prevtblTransactionSub.trnID"
    Flag =3
End
Begin Groups
    Expression ="prevtblAccount.accNo"
    GroupLevel =0
    Expression ="prevtblAccount.caoID"
    GroupLevel =0
    Expression ="prevtblAccount.accName"
    GroupLevel =0
    Expression ="prevtblAccount.accOpenBalance"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="Query8.accName"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
Begin
    State =0
    Left =18
    Top =8
    Right =1898
    Bottom =853
    Left =-1
    Top =-1
    Right =1863
    Bottom =408
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =7
        Right =274
        Bottom =310
        Top =0
        Name ="prevtblAccount"
        Name =""
    End
    Begin
        Left =353
        Top =46
        Right =557
        Bottom =194
        Top =0
        Name ="prevtblTransactionSub"
        Name =""
    End
    Begin
        Left =670
        Top =61
        Right =968
        Bottom =282
        Top =0
        Name ="prevtblTransaction"
        Name =""
    End
End
