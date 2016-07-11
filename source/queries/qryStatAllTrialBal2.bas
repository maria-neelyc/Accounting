dbMemo "SQL" ="SELECT qryChartWithAccounts.coaRef1, qryChartWithAccounts.coaRef2 AS coaRef2, qr"
    "yChartWithAccounts.coaRef3 AS coaRef3, qryChartWithAccounts.coaRef4 AS coaRef4, "
    "qryChartWithAccounts.coaRef5 AS coaRef5, IIf(Sum(qryStatAllTrialBal1.Bal)<>0,Sum"
    "(qryStatAllTrialBal1.Bal),Sum(qryChartWithAccounts.accOpenBalance)) AS Bal, Sum("
    "qryChartWithAccounts.accOpenBalance) AS OpenBal, qryStatAllTrialBal1.CR, qryStat"
    "AllTrialBal1.DR, IIf(Sum(qryStatAllTrialBal1.FBal)<>0,Sum(qryStatAllTrialBal1.FB"
    "al),Sum(qryChartWithAccounts.accOpenBalCur)) AS FBal, Sum(qryChartWithAccounts.a"
    "ccOpenBalCur) AS FOpenBal, Sum(qryStatAllTrialBal1.FCR) AS FCR, Sum(qryStatAllTr"
    "ialBal1.FDR) AS FDR, qryChartWithAccounts.accID AS accID, qryChartWithAccounts.a"
    "ccNo AS accNo, qryChartWithAccounts.accName AS accName, qryChartWithAccounts.acc"
    "Cur\015\012FROM qryChartWithAccounts LEFT JOIN qryStatAllTrialBal1 ON qryChartWi"
    "thAccounts.coaID=qryStatAllTrialBal1.caoID\015\012GROUP BY qryChartWithAccounts."
    "coaRef1, qryChartWithAccounts.coaRef2, qryChartWithAccounts.coaRef3, qryChartWit"
    "hAccounts.coaRef4, qryChartWithAccounts.coaRef5, qryStatAllTrialBal1.CR, qryStat"
    "AllTrialBal1.DR, qryChartWithAccounts.accID, qryChartWithAccounts.accNo, qryChar"
    "tWithAccounts.accName, qryChartWithAccounts.accCur\015\012ORDER BY qryChartWithA"
    "ccounts.coaRef1, qryChartWithAccounts.coaRef2, qryChartWithAccounts.coaRef3, qry"
    "ChartWithAccounts.coaRef4, qryChartWithAccounts.coaRef5;\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="accID"
    End
    Begin
        dbText "Name" ="accNo"
    End
    Begin
        dbText "Name" ="accName"
    End
    Begin
        dbText "Name" ="coaRef2"
    End
    Begin
        dbText "Name" ="coaRef3"
    End
    Begin
        dbText "Name" ="coaRef4"
    End
    Begin
        dbText "Name" ="coaRef5"
    End
    Begin
        dbText "Name" ="OpenBal"
    End
    Begin
        dbText "Name" ="FBal"
    End
    Begin
        dbText "Name" ="FOpenBal"
    End
    Begin
        dbText "Name" ="FCR"
    End
    Begin
        dbText "Name" ="FDR"
    End
End
