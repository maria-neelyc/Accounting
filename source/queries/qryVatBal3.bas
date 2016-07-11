Operation =1
Option =0
Where ="(((tblChart.lvlID)<=[Forms].[frmReport].[txtToStr2]))"
Begin InputTables
    Name ="tblChart"
    Name ="qryVatBal2"
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
    Expression ="qryVatBal2.accID"
    Alias ="Bal"
    Expression ="IIf((IsNumeric(qryVatBal2.CR)),qryVatBal2.Bal,qryVatBal2.OpenBal)"
    Expression ="qryVatBal2.CR"
    Expression ="qryVatBal2.DR"
    Expression ="qryVatBal2.V_CR"
    Expression ="qryVatBal2.V_DR"
    Alias ="OpenBalDR"
    Expression ="IIf(qryVatBal2.OpenBal>=0,qryVatBal2.OpenBal,0)"
    Alias ="OpenBalCR"
    Expression ="IIf(qryVatBal2.OpenBal<0,qryVatBal2.OpenBal,0)"
    Expression ="qryVatBal2.accNo"
    Expression ="qryVatBal2.accName"
    Expression ="qryVatBal2.accType"
    Expression ="tblChart.coaRef1"
    Expression ="tblChart.coaRef2"
    Expression ="tblChart.coaRef3"
    Expression ="tblChart.coaRef4"
    Expression ="tblChart.coaRef5"
    Expression ="tblChart.coaAccType"
    Expression ="qryVatBal2.VatType"
End
Begin Joins
    LeftTable ="tblChart"
    RightTable ="qryVatBal2"
    Expression ="tblChart.coaRef1 = qryVatBal2.coaRef1"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryVatBal2"
    Expression ="tblChart.coaRef2 = qryVatBal2.coaRef2"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryVatBal2"
    Expression ="tblChart.coaRef3 = qryVatBal2.coaRef3"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryVatBal2"
    Expression ="tblChart.coaRef4 = qryVatBal2.coaRef4"
    Flag =1
    LeftTable ="tblChart"
    RightTable ="qryVatBal2"
    Expression ="tblChart.coaRef5 = qryVatBal2.coaRef5"
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
        dbText "Name" ="coaRef"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="coaFRef"
        dbInteger "ColumnWidth" ="3795"
        dbBoolean "ColumnHidden" ="0"
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
    Left =15
    Top =151
    Right =1233
    Bottom =488
    Left =-1
    Top =-1
    Right =1201
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =106
        Top =8
        Name ="tblChart"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =106
        Top =10
        Name ="qryVatBal2"
        Name =""
    End
End
