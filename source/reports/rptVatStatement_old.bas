Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10714
    DatasheetFontHeight =10
    ItemSuffix =65
    Left =1680
    Top =1185
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    Toolbar ="PrintReportEmail"
    RecSrcDt = Begin
        0x97c0d4f96f9ae340
    End
    RecordSource ="qryVatStat"
    Caption ="Vat Statement Report"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x3702000037020000370200003702000000000000da2900005b01000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =238
            FontSize =16
            FontWeight =700
            FontName ="Arial"
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontCharSet =238
            TextFontFamily =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="accNo"
        End
        Begin PageHeader
            Height =2815
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2085
                    Top =1260
                    Height =375
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="accNo"
                    ControlSource ="accNo"

                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3898
                    Top =1260
                    Width =4206
                    Height =375
                    FontSize =12
                    FontWeight =700
                    Name ="accName"
                    ControlSource ="accName"

                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1185
                    Top =2085
                    Width =956
                    Height =227
                    FontSize =10
                    Name ="Label18"
                    Caption ="Entry Date"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =4755
                    Top =2085
                    Width =1247
                    Height =227
                    FontSize =10
                    Name ="Label19"
                    Caption ="Notes"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =2250
                    Top =2085
                    Width =1015
                    Height =227
                    FontSize =10
                    Name ="Label20"
                    Caption ="Reference"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =225
                    Top =2085
                    Width =735
                    Height =240
                    FontSize =10
                    Name ="Label21"
                    Caption ="Voucher"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =3435
                    Top =2085
                    Width =1078
                    Height =227
                    FontSize =10
                    Name ="Label22"
                    Caption ="Description"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =7365
                    Top =2085
                    Width =793
                    Height =227
                    FontSize =10
                    Name ="Label23"
                    Caption ="Debits"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =8610
                    Top =2085
                    Width =737
                    Height =227
                    FontSize =10
                    Name ="Label24"
                    Caption ="Credits"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =161
                    TextAlign =2
                    TextFontFamily =18
                    Left =2370
                    Top =420
                    Width =5925
                    Height =645
                    FontSize =26
                    Name ="Label10"
                    Caption ="VAT"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextFontCharSet =161
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =3120
                    Width =4545
                    Height =388
                    FontSize =15
                    FontWeight =700
                    TabIndex =2
                    Name ="usrName"
                    ControlSource ="=FnIme()"
                    FontName ="Times New Roman"
                    AsianLineBreak =0

                End
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextFontCharSet =161
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =0
                    Left =8552
                    Top =765
                    Width =1964
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="Text12"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"
                    AsianLineBreak =0

                End
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextFontCharSet =161
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =0
                    Left =8295
                    Top =484
                    Width =1739
                    Height =256
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="Text11"
                    ControlSource ="=Now()"
                    Format ="Medium Date"
                    FontName ="Times New Roman"
                    AsianLineBreak =0

                End
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextFontCharSet =161
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =0
                    Left =10059
                    Top =484
                    Width =449
                    Height =256
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    Name ="Text22"
                    ControlSource ="=Time()"
                    Format ="Short Time"
                    FontName ="Times New Roman"
                    AsianLineBreak =0

                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =1144
                    Width =10650
                    BorderColor =4210752
                    Name ="Line69"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =18
                    Left =450
                    Top =1260
                    Width =1485
                    Height =300
                    FontSize =12
                    Name ="Label63"
                    Caption ="Account:"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =2451
                    Width =10605
                    BorderColor =4210752
                    Name ="Line32"
                End
                Begin Label
                    Left =2310
                    Top =2535
                    Width =2440
                    Height =227
                    FontSize =8
                    FontWeight =400
                    Name ="Label37"
                    Caption ="Balance B/F"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7140
                    Top =2535
                    Width =1134
                    TabIndex =6
                    Name ="BalanceBFDR"
                    ControlSource ="=IIf(IsNumeric([Bal]-([CR]-[DR])),IIf([Bal]-([CR]-[DR])<0,Abs([Bal]-([CR]-[DR]))"
                        "),IIf(IsNumeric([OpenBalDR]),IIf([OpenBalDR]=0,\"\",Abs([OpenBalDR]))))"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8385
                    Top =2535
                    Width =1134
                    TabIndex =7
                    Name ="BalanceBFCR"
                    ControlSource ="=IIf(IsNumeric([Bal]-([CR]-[DR])),IIf([Bal]-([CR]-[DR])>=0,[Bal]-([CR]-[DR]),\"\""
                        "),[OpenBalCR])"
                    Format ="Standard"

                End
                Begin Line
                    Left =8325
                    Top =2475
                    Width =0
                    Height =340
                    Name ="Line41"
                End
                Begin Line
                    Left =1129
                    Top =2475
                    Width =0
                    Height =340
                    Name ="Line42"
                End
                Begin Line
                    Left =2209
                    Top =2475
                    Width =0
                    Height =340
                    Name ="Line43"
                End
                Begin Line
                    Left =4714
                    Top =2475
                    Width =0
                    Height =340
                    Name ="Line44"
                End
                Begin Line
                    Left =7080
                    Top =2475
                    Width =0
                    Height =340
                    Name ="Line45"
                End
                Begin Line
                    Top =2775
                    Width =10710
                    Name ="Line46"
                End
                Begin Line
                    Left =3394
                    Top =2475
                    Width =0
                    Height =340
                    Name ="Line47"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5775
                    Top =1695
                    Width =1200
                    Height =270
                    FontSize =10
                    Name ="lblPrdFrom"
                    Caption ="Period From:"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =8205
                    Top =1695
                    Width =1200
                    Height =270
                    FontSize =10
                    Name ="lblPrdTo"
                    Caption ="To:"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7035
                    Top =1695
                    Width =1086
                    TabIndex =8
                    Name ="PeriodFrom"
                    ControlSource ="=Forms!frmReport!txtFromStr2 & ' (' & IIf(Forms!frmReport!txtFromStr2=0,0,MonthN"
                        "ame(Forms!frmReport!txtFromStr2)) & ')'"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9465
                    Top =1695
                    Width =1086
                    TabIndex =9
                    Name ="PeriodTo"
                    ControlSource ="=Forms!frmReport!txtToStr2 & ' (' & IIf(Forms!frmReport!txtToStr2=0,0,MonthName("
                        "Forms!frmReport!txtToStr2)) & ')'"

                End
                Begin Line
                    Left =9525
                    Top =2475
                    Width =0
                    Height =340
                    Name ="Line58"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8265
                    Top =1260
                    Height =375
                    FontSize =12
                    FontWeight =700
                    TabIndex =10
                    Name ="VatType"
                    ControlSource ="VatType"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =347
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1185
                    Top =60
                    Width =906
                    TabIndex =1
                    Name ="trnEntryDate"
                    ControlSource ="trnEntryDate"
                    Format ="Short Date"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2250
                    Top =60
                    Width =1071
                    Name ="trnsNote"
                    ControlSource ="refName"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3390
                    Top =60
                    Width =1251
                    TabIndex =2
                    Name ="refName"
                    ControlSource ="desName"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =-4
                    Top =56
                    Width =1131
                    TabIndex =3
                    Name ="trnInternalRef"
                    ControlSource ="trnInternalRef"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4797
                    Top =56
                    Width =2226
                    ColumnWidth =6315
                    TabIndex =4
                    Name ="desName"
                    ControlSource ="trnsNote"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7140
                    Top =60
                    Width =1134
                    TabIndex =5
                    Name ="txttrnsDebits"
                    ControlSource ="=IIf(IsNumeric([trnsDebits]),IIf([trnsDebits]=0,\"\",[trnsDebits]),\"\")"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8385
                    Top =60
                    Width =1134
                    TabIndex =6
                    Name ="txttrnsCredits"
                    ControlSource ="=IIf(IsNumeric([trnsCredits]),IIf([trnsCredits]=0,\"\",[trnsCredits]),\"\")"
                    Format ="Standard"

                End
                Begin Line
                    Left =8325
                    Width =0
                    Height =340
                    Name ="Line11"
                End
                Begin Line
                    Left =1125
                    Width =0
                    Height =340
                    Name ="Line13"
                End
                Begin Line
                    Left =2205
                    Width =0
                    Height =340
                    Name ="Line14"
                End
                Begin Line
                    Left =4710
                    Width =0
                    Height =340
                    Name ="Line16"
                End
                Begin Line
                    Left =7080
                    Width =0
                    Height =340
                    Name ="Line17"
                End
                Begin Line
                    Top =345
                    Width =10710
                    Name ="Line27"
                End
                Begin Line
                    Left =3390
                    Width =0
                    Height =340
                    Name ="Line30"
                End
                Begin Line
                    Left =9525
                    Width =0
                    Height =340
                    Name ="Line59"
                End
                Begin Label
                    Left =9585
                    Top =60
                    Width =1015
                    Height =240
                    FontSize =8
                    FontWeight =400
                    Name ="lblBalSum"
                End
                Begin Subform
                    Visible = NotDefault
                    OldBorderStyle =0
                    Top =285
                    Width =10710
                    Height =62
                    TabIndex =7
                    Name ="rptVarStatDetails"
                    SourceObject ="Report.rptVarStatDetails"
                    LinkChildFields ="trnsID"
                    LinkMasterFields ="trnsVID"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =585
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7140
                    Top =60
                    Width =1134
                    Name ="TotalDR"
                    ControlSource ="=Sum([trnsDebits])"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8385
                    Top =60
                    Width =1134
                    TabIndex =1
                    Name ="TotalCR"
                    ControlSource ="=Sum([trnsCredits])"
                    Format ="Standard"

                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =6285
                    Top =60
                    Width =737
                    Height =227
                    FontSize =10
                    Name ="Label36"
                    Caption ="Totals:"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7140
                    Top =345
                    Width =1134
                    TabIndex =2
                    Name ="BalanceDB"
                    ControlSource ="=IIf([Bal]<0,Abs([Bal]))"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8385
                    Top =345
                    Width =1134
                    TabIndex =3
                    Name ="BalanceCR"
                    ControlSource ="=IIf([Bal]>=0,[Bal])"
                    Format ="Standard"

                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =5310
                    Top =345
                    Width =1712
                    Height =227
                    FontSize =10
                    Name ="Label50"
                    Caption ="Totals -> Balance:"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin PageFooter
            Height =510
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =390
                    Width =1425
                    Height =268
                    FontSize =9
                    FontWeight =700
                    Name ="Text34"
                    FontName ="Times New Roman"
                    AsianLineBreak =0

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
Private Sub Report_NoData(Cancel As Integer)

MsgBox ("The report is empty.")
Cancel = True

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
'   lblBalSum.Caption = Format(CDbl(IIf(IsNumeric(lblBalSum.Caption), lblBalSum.Caption, 0)) + CDbl(IIf(IsNull(trnsCredits), 0, trnsCredits)) - CDbl(IIf(IsNull(trnsDebits), 0, trnsDebits)), "##########0.00")
End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
'   lblBalSum.Caption = CStr(IIf(IsNull(BalanceBFDR), BalanceBFCR, BalanceBFDR))
End Sub

Private Sub Report_Open(Cancel As Integer)
If IsNull(Forms.frmReport.cboFromStr1) = True Or Forms.frmReport.cboFromStr1 = "" Then
    Forms.frmReport.cboFromStr1.Value = "0"
End If
If IsNull(Forms.frmReport.cboToStr1) = True Or Forms.frmReport.cboToStr1 = "" Then
    Forms.frmReport.cboToStr1.Value = "9999999999"
End If
If IsNull(Forms.frmReport.txtFromDate1) = True Or Forms.frmReport.txtFromDate1 = "" Then
    Forms.frmReport.txtFromDate1.Value = "01-01-1900"
End If
If IsNull(Forms.frmReport.txtToDate1) = True Or Forms.frmReport.txtToDate1 = "" Then
    Forms.frmReport.txtToDate1.Value = Format(Date, "dd-mm-yyyy")
End If
If IsNull(Forms.frmReport.txtFromStr2) = True Or Forms.frmReport.txtFromStr2 = "" Then
    Forms.frmReport.txtFromStr2.Value = 1
End If
If IsNull(Forms.frmReport.txtToStr2) = True Or Forms.frmReport.txtToStr2 = "" Then
    Forms.frmReport.txtToStr2.Value = 12
End If

'show parent transactions
Me.rptVarStatDetails.Visible = Nz(Forms.frmReport.chkBox1, 0)

If Forms.frmReport.txtFromStr2.Value = 0 Then
    PeriodFrom.Visible = False
    lblPrdFrom.Caption = "Period"
    lblPrdTo.Caption = "Up To:"
Else
    PeriodFrom.Visible = True
    lblPrdFrom.Caption = "Period From:"
    lblPrdTo.Caption = "To:"
End If


Select Case Forms.frmReport.grpOptions
    Case 1
        Me.Filter = "Bal = 0"
        Me.FilterOn = True
    Case 2
        Me.Filter = "CR Is Null And DR Is Null"
        Me.FilterOn = True
    Case 3
        Me.Filter = "CR >0 Or DR > 0"
        Me.FilterOn = True
    Case 4
        Me.Filter = ""
        Me.FilterOn = True
End Select
 


End Sub
