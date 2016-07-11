Operation =1
Option =0
Where ="(((qryStatTrialBal3.accNo)>'' And (qryStatTrialBal3.accNo) Between [Forms].[frmR"
    "eport].[txtFromStr1] And [Forms].[frmReport].[txtToStr1]))"
Begin InputTables
    Name ="tblAccount"
    Name ="qryStatementTrns"
    Name ="qryStatTrialBal3"
    Name ="tblCurrency"
End
Begin OutputColumns
    Expression ="qryStatTrialBal3.accNo"
    Expression ="qryStatTrialBal3.accName"
    Expression ="qryStatTrialBal3.Bal"
    Expression ="qryStatementTrns.trnsDocDate"
    Expression ="qryStatementTrns.trnEntryDate"
    Expression ="qryStatementTrns.trnsDebits"
    Expression ="qryStatementTrns.trnsCredits"
    Expression ="qryStatementTrns.trnInternalRef"
    Expression ="qryStatementTrns.refName"
    Expression ="qryStatementTrns.desName"
    Expression ="qryStatementTrns.trnsDate"
    Expression ="qryStatementTrns.yearID"
    Expression ="qryStatementTrns.persID"
    Expression ="qryStatementTrns.trnsNote"
    Expression ="qryStatTrialBal3.CR"
    Expression ="qryStatTrialBal3.DR"
    Expression ="qryStatTrialBal3.OpenBalDR"
    Expression ="qryStatTrialBal3.OpenBalCR"
    Expression ="qryStatementTrns.docNo"
    Expression ="tblCurrency.crnShortName"
    Expression ="tblAccount.accOpenBalCur"
    Expression ="tblCurrency.crnExchangeRate"
End
Begin Joins
    LeftTable ="qryStatementTrns"
    RightTable ="qryStatTrialBal3"
    Expression ="qryStatementTrns.accID = qryStatTrialBal3.accID"
    Flag =3
    LeftTable ="tblAccount"
    RightTable ="qryStatTrialBal3"
    Expression ="tblAccount.accID = qryStatTrialBal3.accID"
    Flag =1
    LeftTable ="tblAccount"
    RightTable ="tblCurrency"
    Expression ="tblAccount.accCur = tblCurrency.crnID"
    Flag =1
End
Begin OrderBy
    Expression ="qryStatTrialBal3.accNo"
    Flag =0
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
    Left =1
    Top =124
    Right =1231
    Bottom =618
    Left =-1
    Top =-1
    Right =1223
    Bottom =326
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =642
        Top =45
        Right =738
        Bottom =152
        Top =8
        Name ="tblAccount"
        Name =""
    End
    Begin
        Left =224
        Top =5
        Right =320
        Bottom =112
        Top =9
        Name ="qryStatementTrns"
        Name =""
    End
    Begin
        Left =380
        Top =9
        Right =476
        Bottom =206
        Top =0
        Name ="qryStatTrialBal3"
        Name =""
    End
    Begin
        Left =761
        Top =46
        Right =857
        Bottom =153
        Top =1
        Name ="tblCurrency"
        Name =""
    End
End
