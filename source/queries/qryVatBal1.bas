Operation =1
Option =0
Where ="(((tblPeriod.perID) Between [Forms].[frmReport].[txtFromStr1] And [Forms].[frmRe"
    "port].[txtToStr1]) AND ((tblAccount.accStatus)=True) AND ((tblTransaction.trnYea"
    "r)=[tblControl].[yearID]))"
Begin InputTables
    Name ="qryVatAcc"
    Name ="tblTransaction"
    Name ="tblPeriodSel"
    Name ="tblPeriod"
    Name ="tblControl"
    Name ="tblAccount"
    Name ="tblTransactionSub"
    Name ="tblTransactionSub"
    Alias ="tblTransactionSub_1"
End
Begin OutputColumns
    Alias ="CR"
    Expression ="Sum(tblTransactionSub.trnsCredits)"
    Alias ="DR"
    Expression ="Sum(tblTransactionSub.trnsDebits)"
    Alias ="Bal"
    Expression ="IIf(Forms.frmReport.txtFromStr1=0,Sum(tblTransactionSub.trnsDebits-tblTransactio"
        "nSub.trnsCredits)+tblAccount.accOpenBalance,Sum(tblTransactionSub.trnsDebits-tbl"
        "TransactionSub.trnsCredits))"
    Alias ="VAmount_Cr"
    Expression ="Sum(tblTransactionSub_1.trnsCredits)"
    Alias ="VAmount_Dr"
    Expression ="Sum(tblTransactionSub_1.trnsDebits)"
    Expression ="tblAccount.accID"
    Expression ="tblAccount.caoID"
    Expression ="tblAccount.accNo"
    Expression ="tblAccount.accName"
End
Begin Joins
    LeftTable ="tblPeriodSel"
    RightTable ="tblPeriod"
    Expression ="tblPeriodSel.perID = tblPeriod.perID"
    Flag =1
    LeftTable ="tblTransaction"
    RightTable ="tblPeriodSel"
    Expression ="tblTransaction.persID = tblPeriodSel.persID"
    Flag =1
    LeftTable ="tblTransaction"
    RightTable ="tblControl"
    Expression ="tblTransaction.trnYear = tblControl.yearID"
    Flag =1
    LeftTable ="qryVatAcc"
    RightTable ="tblAccount"
    Expression ="qryVatAcc.accNo = tblAccount.accNo"
    Flag =1
    LeftTable ="tblTransactionSub"
    RightTable ="tblTransactionSub_1"
    Expression ="tblTransactionSub.trnsVID = tblTransactionSub_1.trnsID"
    Flag =2
    LeftTable ="tblAccount"
    RightTable ="tblTransactionSub"
    Expression ="tblAccount.accID = tblTransactionSub.accID"
    Flag =1
    LeftTable ="tblTransaction"
    RightTable ="tblTransactionSub"
    Expression ="tblTransaction.trnID = tblTransactionSub.trnID"
    Flag =1
End
Begin Groups
    Expression ="tblAccount.accID"
    GroupLevel =0
    Expression ="tblAccount.caoID"
    GroupLevel =0
    Expression ="tblAccount.accNo"
    GroupLevel =0
    Expression ="tblAccount.accName"
    GroupLevel =0
    Expression ="tblAccount.accOpenBalance"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Bal"
    End
    Begin
        dbText "Name" ="CR"
    End
    Begin
        dbText "Name" ="DR"
    End
End
Begin
    State =0
    Left =1
    Top =87
    Right =1615
    Bottom =528
    Left =-1
    Top =-1
    Right =1597
    Bottom =248
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =842
        Top =6
        Right =938
        Bottom =106
        Top =0
        Name ="qryVatAcc"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =106
        Top =0
        Name ="tblTransaction"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =106
        Top =0
        Name ="tblPeriodSel"
        Name =""
    End
    Begin
        Left =574
        Top =6
        Right =670
        Bottom =106
        Top =0
        Name ="tblPeriod"
        Name =""
    End
    Begin
        Left =708
        Top =6
        Right =804
        Bottom =106
        Top =0
        Name ="tblControl"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =106
        Top =0
        Name ="tblAccount"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =411
        Bottom =202
        Top =0
        Name ="tblTransactionSub"
        Name =""
    End
    Begin
        Left =986
        Top =7
        Right =1425
        Bottom =203
        Top =12
        Name ="tblTransactionSub_1"
        Name =""
    End
End
