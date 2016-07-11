Operation =1
Option =0
Where ="(((tblChart.lvlID)=5))"
Begin InputTables
    Name ="tblChart"
    Name ="qryStatTrialBal2"
End
Begin OutputColumns
    Expression ="tblChart.coaID"
    Expression ="tblChart.coaNo"
    Expression ="tblChart.coaName"
    Alias ="coaFRef"
    Expression ="Right(Chr(48) & tblChart.coaRef1,2) & \"-\" & Right(Chr(48) & tblChart.coaRef2,2"
        ") & \"-\" & Right(Chr(48) & tblChart.coaRef3,2) & \"-\" & Right(Chr(48) & tblCha"
        "rt.coaRef4,2) & \"-\" & Right(Chr(48) & Chr(48) & tblChart.coaRef5,3)"
    Alias ="coaRef"
    Expression ="Choose(tblChart.lvlID,Left(coaFRef,2),Left(coaFRef,5),Left(coaFRef,8),Left(coaFR"
        "ef,11),coaFRef)"
    Expression ="tblChart.lvlID"
    Expression ="qryStatTrialBal2.accID"
    Alias ="Bal"
    Expression ="IIf((IsNumeric(qryStatTrialBal2.CR)),qryStatTrialBal2.Bal,qryStatTrialBal2.OpenB"
        "al)"
    Expression ="qryStatTrialBal2.CR"
    Expression ="qryStatTrialBal2.DR"
    Alias ="OpenBalDR"
    Expression ="IIf(qryStatTrialBal2.OpenBal>=0,qryStatTrialBal2.OpenBal,0)"
    Alias ="OpenBalCR"
    Expression ="IIf(qryStatTrialBal2.OpenBal<0,qryStatTrialBal2.OpenBal,0)"
    Expression ="qryStatTrialBal2.accNo"
    Expression ="qryStatTrialBal2.accName"
End
Begin Joins
    LeftTable ="tblChart"
    RightTable ="qryStatTrialBal2"
    Expression ="tblChart.coaRef5=qryStatTrialBal2.coaRef5"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryStatTrialBal2"
    Expression ="tblChart.coaRef4=qryStatTrialBal2.coaRef4"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryStatTrialBal2"
    Expression ="tblChart.coaRef3=qryStatTrialBal2.coaRef3"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryStatTrialBal2"
    Expression ="tblChart.coaRef2=qryStatTrialBal2.coaRef2"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryStatTrialBal2"
    Expression ="tblChart.coaRef1=qryStatTrialBal2.coaRef1"
    Flag =1
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
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
Begin
    Begin
        dbText "Name" ="coaFRef"
    End
    Begin
        dbText "Name" ="coaRef"
    End
    Begin
        dbText "Name" ="OpenBalDR"
    End
    Begin
        dbText "Name" ="OpenBalCR"
    End
    Begin
        dbText "Name" ="Bal"
    End
End
Begin
    State =0
    Left =28
    Top =227
    Right =1246
    Bottom =539
    Left =-1
    Top =-1
    Right =1211
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
        Right =268
        Bottom =113
        Top =8
        Name ="qryStatTrialBal2"
        Name =""
    End
End
