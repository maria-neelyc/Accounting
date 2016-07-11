Operation =1
Option =0
Where ="(((Year([trnsDate])) In (SELECT yearID FROM prevtblControl)) AND ((prevtblAccoun"
    "t.accType)=2))"
Begin InputTables
    Name ="prevtblAccount"
    Name ="prevtblTransactionSub"
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
    Bottom =713
    Left =-1
    Top =-1
    Right =1859
    Bottom =267
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =7
        Right =210
        Bottom =211
        Top =0
        Name ="prevtblAccount"
        Name =""
    End
    Begin
        Left =373
        Top =-7
        Right =538
        Bottom =204
        Top =0
        Name ="prevtblTransactionSub"
        Name =""
    End
End
