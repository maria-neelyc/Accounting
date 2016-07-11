Operation =1
Option =0
Where ="(((tblChart.lvlID)=5))"
Begin InputTables
    Name ="tblChart"
    Name ="qryStatForTrialBal2"
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
    Expression ="qryStatForTrialBal2.accID"
    Alias ="Bal"
    Expression ="IIf((IsNumeric(qryStatForTrialBal2.CR)),qryStatForTrialBal2.Bal,qryStatForTrialB"
        "al2.OpenBal)"
    Expression ="qryStatForTrialBal2.CR"
    Expression ="qryStatForTrialBal2.DR"
    Alias ="OpenBalDR"
    Expression ="IIf(qryStatForTrialBal2.OpenBal>=0,qryStatForTrialBal2.OpenBal,0)"
    Alias ="OpenBalCR"
    Expression ="IIf(qryStatForTrialBal2.OpenBal<0,qryStatForTrialBal2.OpenBal,0)"
    Expression ="qryStatForTrialBal2.accNo"
    Expression ="qryStatForTrialBal2.accName"
    Expression ="qryStatForTrialBal2.accCur"
End
Begin Joins
    LeftTable ="tblChart"
    RightTable ="qryStatForTrialBal2"
    Expression ="tblChart.coaRef1=qryStatForTrialBal2.coaRef1"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryStatForTrialBal2"
    Expression ="tblChart.coaRef2=qryStatForTrialBal2.coaRef2"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryStatForTrialBal2"
    Expression ="tblChart.coaRef3=qryStatForTrialBal2.coaRef3"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryStatForTrialBal2"
    Expression ="tblChart.coaRef4=qryStatForTrialBal2.coaRef4"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryStatForTrialBal2"
    Expression ="tblChart.coaRef5=qryStatForTrialBal2.coaRef5"
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
    Right =1207
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
        Top =0
        Name ="qryStatForTrialBal2"
        Name =""
    End
End
