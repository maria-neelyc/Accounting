Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =8267
    DatasheetFontHeight =10
    ItemSuffix =274
    Left =4380
    Top =2100
    Right =12645
    Bottom =8955
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x680150711b99e340
    End
    RecordSource ="SELECT tblControl.*, sysMyCo.smycoName, sysMyCo.smycoPath, sysMyCo.smycoYear FRO"
        "M tblControl INNER JOIN sysMyCo ON tblControl.ctrDBID=sysMyCo.smycoID; "
    Caption ="Control File"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa4050000a2050000a2050000a205000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =755
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =465
                    Top =180
                    Width =6990
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="lblComent"
                    Caption ="System Parameters and Company Informations"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5535
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Left =45
                    Top =45
                    Width =8163
                    Height =5490
                    Name ="tabAcc"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =180
                            Top =450
                            Width =7890
                            Height =4950
                            Name ="pgeGen"
                            Caption ="Account Details"
                            LayoutCachedLeft =180
                            LayoutCachedTop =450
                            LayoutCachedWidth =8070
                            LayoutCachedHeight =5400
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =4988
                                    Top =452
                                    Width =180
                                    Height =255
                                    Name ="ctrID"
                                    ControlSource ="ctrID"

                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =161
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2150
                                    Top =796
                                    Width =3330
                                    Height =255
                                    TabIndex =1
                                    Name ="ctrName"
                                    ControlSource ="ctrName"
                                    FontName ="Arial Greek"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =710
                                            Top =796
                                            Width =1380
                                            Height =255
                                            Name ="bnkaABA_Label"
                                            Caption ="Company Name"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =161
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2150
                                    Top =1126
                                    Width =3330
                                    Height =255
                                    TabIndex =2
                                    Name ="ctrDate"
                                    ControlSource ="ctrDate"
                                    FontName ="Arial Greek"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =710
                                            Top =1126
                                            Width =1380
                                            Height =255
                                            Name ="bnkaBLZ_Label"
                                            Caption ="Date"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =161
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2150
                                    Top =1434
                                    Width =3330
                                    Height =255
                                    TabIndex =3
                                    Name ="ctrPageCounter"
                                    ControlSource ="ctrPageCounter"
                                    FontName ="Arial Greek"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =710
                                            Top =1434
                                            Width =1380
                                            Height =255
                                            Name ="Label84"
                                            Caption ="Page No"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =1134
                                    Left =2812
                                    Top =2520
                                    Width =1740
                                    Height =270
                                    TabIndex =4
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                                    Name ="persID"
                                    ControlSource ="persID"
                                    RowSourceType ="Table/Query"
                                    RowSource ="tblPeriodSel"
                                    ColumnWidths ="0;1134"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =988
                                            Top =2524
                                            Width =1080
                                            Height =255
                                            Name ="Label35"
                                            Caption ="Current Period"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =2157
                                    Top =3379
                                    Width =3360
                                    Height =270
                                    TabIndex =5
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="ctrLYPLA"
                                    ControlSource ="ctrLYPLA"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    OnDblClick ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =703
                                            Top =3380
                                            Width =1395
                                            Height =255
                                            Name ="lblLYPLA"
                                            Caption ="Last Profit &&  Loose"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    TextAlign =2
                                    BackStyle =0
                                    Left =2158
                                    Top =2539
                                    Width =507
                                    Height =270
                                    TabIndex =6
                                    Name ="TXTpersID"
                                    ControlSource ="persID"

                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =2907
                                    Top =2129
                                    Width =1296
                                    Height =351
                                    TabIndex =7
                                    Name ="btnCloseYear"
                                    Caption ="Close Year >>"
                                    OnClick ="[Event Procedure]"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2151
                                    Top =2189
                                    Width =705
                                    TabIndex =8
                                    Name ="yearID"
                                    ControlSource ="yearID"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1466
                                            Top =2213
                                            Width =615
                                            Height =240
                                            Name ="Label97"
                                            Caption ="Year"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =4793
                                    Top =1998
                                    Width =510
                                    TabIndex =9
                                    Name ="smycoPath"
                                    ControlSource ="smycoPath"

                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =2157
                                    Top =3720
                                    Width =3360
                                    Height =270
                                    TabIndex =10
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="ctrControlAcc"
                                    ControlSource ="ctrControlAcc"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount OR"
                                        "DER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =570
                                            Top =3720
                                            Width =1530
                                            Height =255
                                            Name ="lblControlAcc"
                                            Caption ="Control Account"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =2145
                                    Top =2865
                                    Width =2001
                                    Height =351
                                    TabIndex =11
                                    Name ="cmdUpdtBal"
                                    Caption ="Update Current Balances"
                                    OnClick ="[Event Procedure]"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =4275
                                    Top =2865
                                    Width =2001
                                    Height =351
                                    TabIndex =12
                                    Name ="cmdUpdChart"
                                    Caption ="Update Chart of Accounts"
                                    OnClick ="[Event Procedure]"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =180
                            Top =450
                            Width =7896
                            Height =4950
                            Name ="pgePeriod"
                            Caption ="Current Balance"
                            LayoutCachedLeft =180
                            LayoutCachedTop =450
                            LayoutCachedWidth =8076
                            LayoutCachedHeight =5400
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =757
                                    Top =583
                                    Width =4995
                                    Height =4590
                                    Name ="Reference File"
                                    SourceObject ="Form.frmPeriod"
                                    EventProcPrefix ="Reference_File"

                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5811
                                    Top =4275
                                    Width =2265
                                    Height =900
                                    Name ="Label106"
                                    Caption ="NOTICE:\015\012Any change in \"Period Section will direclt take effect in calcul"
                                        "ation.\015\012"
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =590
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =70
                    Left =210
                    Top =90
                    Width =366
                    Name ="cmdFirst"
                    Caption ="&F"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada44dadad1dadaadad44adad11adaddada44dad111dada ,
                        0xadad44ad1111adaddada44d11111dadaadad44ad1111adaddada44dad111dada ,
                        0xadad44adad11adaddada44dadad1dadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="First Record (Alt+F)"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =76
                    Left =1470
                    Top =90
                    Width =366
                    TabIndex =1
                    Name ="cmdLast"
                    Caption ="&L"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadad1dadad44adaadada11dada44daddadad111dad44ada ,
                        0xadada1111da44daddadad11111d44adaadada1111da44daddadad111dad44ada ,
                        0xadada11dada44daddadad1dadad44adaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Last Record (Alt+L)"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    Left =1050
                    Top =90
                    Width =366
                    TabIndex =2
                    Name ="cmdNext"
                    Caption ="&N"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadada1adadadadaadadad11adadadaddadada111adadada ,
                        0xadadad1111adadaddadada11111adadaadadad1111adadaddadada111adadada ,
                        0xadadad11adadadaddadada1adadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record (Alt+N)"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =80
                    Left =630
                    Top =90
                    Width =366
                    TabIndex =3
                    Name ="cmdPrevious"
                    Caption ="&P"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadad1dadadaadadadad11adadaddadadad111dadada ,
                        0xadadad1111adadaddadad11111dadadaadadad1111adadaddadadad111dadada ,
                        0xadadadad11adadaddadadadad1dadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record (Alt+P)"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    AccessKey =65
                    Left =4500
                    Top =90
                    Width =1080
                    TabIndex =4
                    Name ="cmdAdd"
                    Caption ="&Add New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add New Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    AccessKey =68
                    Left =5700
                    Top =90
                    Width =1080
                    TabIndex =5
                    Name ="cmdDel"
                    Caption ="&Delete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    Left =6900
                    Top =90
                    Width =1080
                    TabIndex =6
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close Current Form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnCloseYear_Click()
Dim strNewDBPath As String
Dim Db As Database
Dim rs As DAO.Recordset
Dim strSQL As String, dblBalance As Double

Const cERR_USERCANCEL = vbObjectError + 1000

strNewDBPath = CreateNewMDBFile

If fRefreshLinks(smycoPath, True) Then

    DoCmd.SetWarnings False
    
    'Updates local control table
    strSQL = "INSERT INTO sysMyCo (smycoName, smycoPath, smycoYear) " & _
             "VALUES ('" & Me.smycoName & "', '" & strNewDBPath & "', " & Me.smycoYear + 1 & ")"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    
    'Sets flag to closed year
    strSQL = "UPDATE tblControl " & _
                 "SET ctrClosed = True"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    
    'Updates company control table
    strSQL = "SELECT smycoID FROM sysMyCo WHERE smycoPath = '" & strNewDBPath & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    rs.MoveFirst
    
    strSQL = "UPDATE nexttblControl " & _
             "SET yearID = " & Me.Recordset.Fields("yearID") + 1 & ", " & _
             "persID = 1, " & _
             "ctrDBID = " & rs.Fields("smycoID")
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
   
    'Clean transactions
    strSQL = "DELETE * FROM nexttblTransaction"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    
    strSQL = "DELETE * FROM nexttblTransactionSub"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)

    'Updates balances

    
    sUpdtBalances
    
    'Check if balances match
    
    strSQL = "SELECT Sum(a.Bal) AS currBal, b.nextBal FROM (" & _
             "SELECT (Sum(IIf([trnsSign]='C',[trnsAmount],-1*[trnsAmount]))) AS Bal " & _
             "FROM (tblAccount INNER JOIN (tblTransactionSub INNER JOIN tblTransaction ON tblTransactionSub.trnID=tblTransaction.trnID) ON tblAccount.accID=tblTransactionSub.accID) INNER JOIN nexttblAccount ON tblAccount.accID=nexttblAccount.accID " & _
             "WHERE (((tblTransaction.trnYear) In (SELECT yearID FROM tblControl))) " & _
             "UNION ALL " & _
             "SELECT  Sum(tblAccount.accOpenBalance) AS Bal " & _
             "FROM tblAccount) a, (SELECT Sum(nexttblAccount.accOpenBalance) As nextBal " & _
             "FROM nexttblAccount) b " & _
             "GROUP BY b.nextBal " & _
             "HAVING Sum(a.Bal) <> b.nextBal"
             
    Set Db = CurrentDb()
    Set rs = Db.OpenRecordset(strSQL)
    If rs.RecordCount Then
        MsgBox FnIsErr(7005), vbExclamation
    Else
        If MsgBox("New year has been succesfully created. Do you want to open it?", _
                vbQuestion + vbYesNo, "Please confirm...") = vbNo Then
            Me.cmdExit.SetFocus
            Me.btnCloseYear.Enabled = False
        Else
            If fRefreshLinks(strNewDBPath, True) Then
                DoCmd.Close acForm, "frmControl"
                DoCmd.OpenForm "frmControl"
            Else
                MsgBox FnIsErr(7006), vbExclamation
            End If
        End If
    End If
    
    rs.Close
    Db.Close
    Set rs = Nothing
    Set Db = Nothing
    DoCmd.SetWarnings True
End If
End Sub

Private Sub cboCRN_DR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboCshWithdrwDR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboDeptsCR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboDRN_CR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboExpDR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboInvCR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboPaymtCR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboRcptDR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cmdCommit_Click()
'Dim bolSuccess As Boolean
'
'Me.frmTransactionLock.Requery
'bolSuccess = fsSequence(0, 0)

End Sub

Private Sub cmdFirst_Click()
On Error GoTo Err_cmdFirst_Click
    DoCmd.GoToRecord , , acFirst
Exit_cmdFirst_Click:
    Exit Sub
Err_cmdFirst_Click:
    If Err.Number = 13 Then GoTo Exit_cmdFirst_Click
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdFirst_Click
End Sub
Private Sub cmdLast_Click()
On Error GoTo Err_cmdLast_Click
    DoCmd.GoToRecord , , acLast
Exit_cmdLast_Click:
    Exit Sub
Err_cmdLast_Click:
    If Err.Number = 13 Then GoTo Exit_cmdLast_Click
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdLast_Click
End Sub
Private Sub cmdNext_Click()
On Error GoTo Err_cmdNext_Click
    DoCmd.GoToRecord , , acNext
Exit_cmdNext_Click:
    Exit Sub
Err_cmdNext_Click:
    Select Case Err.Number
    Case 13
    Case 2105
        Beep
    Case Else
        MsgBox FnIsErr(Err.Number), vbExclamation
    End Select
    Resume Exit_cmdNext_Click
End Sub
Private Sub cmdPrevious_Click()
On Error GoTo Err_cmdPrevious_Click
    DoCmd.GoToRecord , , acPrevious
Exit_cmdPrevious_Click:
    Exit Sub
Err_cmdPrevious_Click:
    Select Case Err.Number
    Case 13
    Case 2105
        Beep
    Case Else
        MsgBox FnIsErr(Err.Number), vbExclamation
    End Select
    Resume Exit_cmdPrevious_Click
End Sub
Private Sub cmdAdd_Click()
On Error GoTo Err_cmdAdd_Click
  Me.AllowAdditions = True
  DoCmd.GoToRecord , , acNewRec
  frmRef1.SetFocus
Exit_cmdAdd_Click:
    Exit Sub
Err_cmdAdd_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdAdd_Click
End Sub
Private Sub cmdDel_Click()
On Error GoTo Err_cmdDel_Click
    DoCmd.SetWarnings False
    If MsgBox("Delete Record? YES or NO?", vbYesNo, "Warning!") = vbYes Then
        DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
        DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
        frmRef1.SetFocus
    Else
        Exit Sub
    End If
    DoCmd.SetWarnings True
Exit_cmdDel_Click:
    Exit Sub
Err_cmdDel_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdDel_Click
End Sub
Private Sub cmdExit_Click()
If IsNoData(Me.OpenArgs) Then
    DoCmd.OpenForm "frmMain"
    DoCmd.Close acForm, Me.name
Else
    Select Case Me.OpenArgs
        Case "frmTransactionSub"
            DoCmd.Close acForm, "frmAccount"
            Screen.ActiveForm.Refresh
        Case Else
    End Select
End If
End Sub

'Private Sub cmdStartRge_Click()
'Dim strPages As String
'Dim i As Integer
'Dim bolSuccess As Boolean
'
'If IsNull(Me.txtFromPage) Then
'    MsgBox ("Fill begining range for pages.")
'    Exit Sub
'End If
'
'If IsNull(Me.txtToPage) Then
'    MsgBox ("Fill ending range for pages.")
'    Exit Sub
'End If
'
'Dim SQL As String
'
'SQL = "UPDATE tblTransaction " & _
'      "SET tblTransaction.trnLock = True " & _
'      "WHERE tblTransaction.trnPageCounter Between " & Me.txtFromPage & " AND " & Me.txtToPage & _
'      " AND tblTransaction.persID <= " & Me.txtUpToPrd
'
'DoCmd.SetWarnings False
'DoCmd.RunSQL SQL
'DoCmd.SetWarnings True
'
'bolSuccess = fsSequence(CInt(Me.txtFromPage), CInt(Me.txtToPage))
'Me.frmTransactionLock.Requery
'End Sub

Private Sub cmdUpdChart_Click()
    Dim rs As Recordset
    
    
    strSQL = "SELECT coaRef1, coaRef2, coaRef3, coaRef4 " & _
             "FROM tblChart " & _
             "WHERE lvlID = 5"
             
    Set rs = CurrentDb.OpenRecordset(strSQL)
    rs.MoveFirst
    Do Until rs.EOF
        Call ChartAccType(rs.Fields("coaRef1"), rs.Fields("coaRef2"), rs.Fields("coaRef3"), rs.Fields("coaRef4"))
        rs.MoveNext
    Loop
    DoCmd.SetWarnings False
    strSQL = "UPDATE tblChart SET coaAccType = 0 WHERE IsNull(coaAccType)"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    DoCmd.SetWarnings True
End Sub

Private Sub cmdUpdtBal_Click()
If fsPrevYearExist Then
    'updates opening balances for the next year based on current db year transactions
    'sUpdtBalances
    
    'updates opening balances for the current db year based on previous year transactions
    sUpdtBalancesCur
End If

'updates opening balances for the next year based on current db year transactions
'If fsClosedYear Then
'    If fsNextYearHasChanged Then
'        MsgBox FnIsErr(7009), vbExclamation
'    End If
'    sUpdtBalances
    
'End If

End Sub

Private Sub ctrControlAcc_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub ctrLYPLA_DblClick(Cancel As Integer)
DoCmd.OpenForm "frmAccount", , , , , , Me.Form.name
Me.ctrLYPLA.Requery
End Sub

Private Sub ctrLYPLA_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
Dim strMsg As String
        strMsg = FnIsErr(DataErr)
        Response = acDataErrContinue
        If IsNoData(strMsg) = False Then
            MsgBox strMsg, vbExclamation, "Error!"
        Else
        End If
End Sub


Private Sub Form_Load()
If Me.Recordset.Fields("ctrClosed") = True Then
'    Me.btnCloseYear.Enabled = False
End If
End Sub

'Private Sub frmLockPages_Click()
'If Me.frmLockPages = 1 Then
'    Me.txtFromPage.Enabled = True
'    Me.txtToPage.Enabled = True
'    Me.txtUpToPrd.Enabled = True
'    Me.cmdStartRge.Enabled = True
'    Me.txtFromPage.SetFocus
'
'    Me.frmLockPages.Requery
'    Me.frmTransactionLock.Enabled = False
'    Me.frmTransactionLock.Form.trnEntryDate.ForeColor = 8421504
'    Me.frmTransactionLock.Form.trnInternalRef.ForeColor = 8421504
'    Me.frmTransactionLock.Form.trnPageCounter.ForeColor = 8421504
'    Me.frmTransactionLock.Form.persID.ForeColor = 8421504
'    Me.frmTransactionLock.Form.trnYear.ForeColor = 8421504
'    Me.frmTransactionLock.Form.trnEntryDate.BorderColor = 8421504
'    Me.frmTransactionLock.Form.trnInternalRef.BorderColor = 8421504
'    Me.frmTransactionLock.Form.trnPageCounter.BorderColor = 8421504
'    Me.frmTransactionLock.Form.persID.BorderColor = 8421504
'    Me.frmTransactionLock.Form.trnYear.BorderColor = 8421504
'    Me.cmdCommit.Enabled = False
'    If Me.frmTransactionLock.Form.Recordset.RecordCount > 0 Then
'        Me.frmTransactionLock.Form.trnLock = 0
'    End If
'Else
'    Me.txtFromPage.Enabled = False
'    Me.txtToPage.Enabled = False
'    Me.txtUpToPrd.Enabled = False
'    Me.cmdStartRge.Enabled = False
'
'    Me.frmTransactionLock.Enabled = True
'    Me.frmTransactionLock.Form.trnEntryDate.ForeColor = 0
'    Me.frmTransactionLock.Form.trnInternalRef.ForeColor = 0
'    Me.frmTransactionLock.Form.trnPageCounter.ForeColor = 0
'    Me.frmTransactionLock.Form.persID.ForeColor = 0
'    Me.frmTransactionLock.Form.trnYear.ForeColor = 0
'    Me.frmTransactionLock.Form.trnEntryDate.BorderColor = 0
'    Me.frmTransactionLock.Form.trnInternalRef.BorderColor = 0
'    Me.frmTransactionLock.Form.trnPageCounter.BorderColor = 0
'    Me.frmTransactionLock.Form.persID.BorderColor = 0
'    Me.frmTransactionLock.Form.trnYear.BorderColor = 0
'    Me.cmdCommit.Enabled = True
'End If
'End Sub


Private Sub FnAllocAC(accID As Integer)
    
    If accID = 0 Then Exit Sub
    
    Dim rs As Recordset
    Dim strSQL As String

    strSQL = "SELECT nz(tblControl.ctrLYPLA,0) AS ctrLYPLA, nz(tblControl.ctrControlAcc,0) AS ctrControlAcc, nz(tblControl.ctrInvCR,0) AS ctrInvCR, nz(tblControl.ctrCRN_DR, 0) AS ctrCRN_DR, nz(tblControl.ctrDRN_CR, 0) AS ctrDRN_CR, nz(tblControl.ctrRcptDR,0) AS ctrRcptDR, nz(tblControl.ctrExpDR,0) AS ctrExpDR, nz(tblControl.ctrPaymtCR,0) AS ctrPaymtCR, nz(tblControl.ctrDeptsCR,0) AS ctrDeptsCR, nz(tblControl.ctrCshWithdrwDR,0) AS ctrCshWithdrwDR " & _
             "FROM tblControl;"
    
    Set rs = CurrentDb.OpenRecordset(strSQL)
    rs.MoveFirst
    If Forms.frmControl.ActiveControl.name <> "ctrLYPLA" _
    And (rs!ctrLYPLA = accID _
    Or Forms.frmControl.ctrLYPLA = accID) Then
        MsgBox FnIsErr(7011, Forms.frmControl.lblLYPLA.Caption), vbExclamation
        DoCmd.CancelEvent
        Exit Sub
    End If

'Checking if selected account was already selected somewhere else
    If Forms.frmControl.ActiveControl.name <> "ctrControlAcc" _
    And (rs!ctrControlAcc = accID _
    Or Forms.frmControl.ctrControlAcc = accID) Then
        MsgBox FnIsErr(7011, Forms.frmControl.lblControlAcc.Caption), vbExclamation
        DoCmd.CancelEvent
        Exit Sub
    End If

'    If Forms.frmControl.ActiveControl.Name <> "cboInvCR" _
'    And (rs!ctrInvCR = accID _
'    Or Forms.frmControl.cboInvCR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblInvCR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboCRN_DR" _
'    And (rs!ctrCRN_DR = accID _
'    Or Forms.frmControl.cboCRN_DR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblCRN_DR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboDRN_CR" _
'    And (rs!ctrDRN_CR = accID _
'    Or Forms.frmControl.cboDRN_CR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblDRN_CR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboRcptDR" _
'    And (rs!ctrRcptDR = accID _
'    Or Forms.frmControl.cboRcptDR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblRcptDR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboExpDR" _
'    And (rs!ctrExpDR = accID _
'    Or Forms.frmControl.cboExpDR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblExpDR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboPaymtCR" _
'    And (rs!ctrPaymtCR = accID _
'    Or Forms.frmControl.cboPaymtCR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblPaymtCR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboDeptsCR" _
'    And (rs!ctrDeptsCR = accID _
'    Or Forms.frmControl.cboDeptsCR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblDeptsCR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboCshWithdrwDR" _
'    And (rs!ctrCshWithdrwDR = accID _
'    Or Forms.frmControl.cboCshWithdrwDR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblCshWithdrwDR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If

End Sub
