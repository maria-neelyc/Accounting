Operation =1
Option =0
Where ="(((prevtblTransaction.trnYear) In (SELECT yearID FROM prevtblControl)) AND ((pre"
    "vtblAccount.accType)=1))"
Begin InputTables
    Name ="prevtblControl"
    Name ="prevtblAccount"
    Name ="prevtblTransactionSub"
    Name ="prevtblTransaction"
    Name ="tblAccount"
End
Begin OutputColumns
    Alias ="Bal"
    Expression ="Sum(IIf([trnsSign]='C',[trnsAmount],-1*[trnsAmount]))"
    Expression ="prevtblControl.ctrLYPLA"
End
Begin Joins
    LeftTable ="prevtblAccount"
    RightTable ="tblAccount"
    Expression ="prevtblAccount.accID = tblAccount.accID"
    Flag =1
    LeftTable ="prevtblAccount"
    RightTable ="prevtblTransactionSub"
    Expression ="prevtblAccount.accID = prevtblTransactionSub.accID"
    Flag =1
    LeftTable ="prevtblTransaction"
    RightTable ="prevtblTransactionSub"
    Expression ="prevtblTransaction.trnID = prevtblTransactionSub.trnID"
    Flag =1
End
Begin Groups
    Expression ="prevtblControl.ctrLYPLA"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Bal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="prevtblControl.ctrLYPLA"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =894
    Bottom =656
    Left =-1
    Top =-1
    Right =870
    Bottom =314
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="prevtblControl"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="prevtblAccount"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="prevtblTransactionSub"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="prevtblTransaction"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="tblAccount"
        Name =""
    End
End
