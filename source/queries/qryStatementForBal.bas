Operation =1
Option =0
Where ="(((qryStatForTrialBal3.accNo)>'' And (qryStatForTrialBal3.accNo) Between Forms.f"
    "rmReport.cboFromStr1 And Forms.frmReport.cboToStr1) And ((qryStatForTrialBal3.ac"
    "cCur)>0))"
Begin InputTables
    Name ="qryStatementForTrns"
    Name ="qryStatForTrialBal3"
End
Begin OutputColumns
    Expression ="qryStatementForTrns.trnsID"
    Expression ="qryStatForTrialBal3.accNo"
    Expression ="qryStatForTrialBal3.accName"
    Expression ="qryStatForTrialBal3.Bal"
    Expression ="qryStatementForTrns.trnsDocDate"
    Expression ="qryStatementForTrns.trnEntryDate"
    Expression ="qryStatementForTrns.trnsDebits"
    Expression ="qryStatementForTrns.trnsCredits"
    Expression ="qryStatementForTrns.trnInternalRef"
    Expression ="qryStatementForTrns.refName"
    Expression ="qryStatementForTrns.desName"
    Expression ="qryStatementForTrns.trnsDate"
    Expression ="qryStatementForTrns.yearID"
    Expression ="qryStatementForTrns.persID"
    Expression ="qryStatementForTrns.trnsNote"
    Expression ="qryStatForTrialBal3.CR"
    Expression ="qryStatForTrialBal3.DR"
    Expression ="qryStatForTrialBal3.OpenBalDR"
    Expression ="qryStatForTrialBal3.OpenBalCR"
    Expression ="qryStatementForTrns.docNo"
    Expression ="qryStatForTrialBal3.accCur"
End
Begin Joins
    LeftTable ="qryStatementForTrns"
    RightTable ="qryStatForTrialBal3"
    Expression ="qryStatementForTrns.accID=qryStatForTrialBal3.accID"
    Flag =3
End
Begin OrderBy
    Expression ="qryStatForTrialBal3.accNo"
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
        Left =38
        Top =6
        Right =194
        Bottom =113
        Top =0
        Name ="qryStatementForTrns"
        Name =""
    End
    Begin
        Left =280
        Top =8
        Right =444
        Bottom =115
        Top =0
        Name ="qryStatForTrialBal3"
        Name =""
    End
End
