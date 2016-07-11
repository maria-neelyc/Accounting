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
    Width =10710
    DatasheetFontHeight =10
    ItemSuffix =74
    Left =312
    Top =168
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    Toolbar ="PrintReportEmail"
    RecSrcDt = Begin
        0x022d22508c8fe440
    End
    RecordSource ="qryVatStat"
    Caption ="Statement Report"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x3702000037020000370200003702000000000000d62900005901000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =238
            FontSize =16
            FontWeight =700
            FontName ="Arial"
        End
        Begin Line
            Width =1701
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontCharSet =238
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin Subform
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
        Begin BreakLevel
            ControlSource ="trnsDocDate"
        End
        Begin BreakLevel
            ControlSource ="trnsID"
        End
        Begin PageHeader
            Height =2770
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =161
                    TextAlign =2
                    TextFontFamily =18
                    Left =2325
                    Top =420
                    Width =5925
                    Height =645
                    FontSize =26
                    Name ="Label10"
                    Caption ="V.A.T. Statement"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextFontCharSet =161
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =2325
                    Width =5895
                    Height =388
                    FontSize =15
                    FontWeight =700
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
                    Left =8507
                    Top =765
                    Width =1964
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
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
                    Left =8250
                    Top =484
                    Width =1739
                    Height =256
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
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
                    Left =10014
                    Top =484
                    Width =449
                    Height =256
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="Text22"
                    ControlSource ="=Time()"
                    Format ="Short Time"
                    FontName ="Times New Roman"
                    AsianLineBreak =0
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =60
                    Top =1144
                    Width =10545
                    Height =30
                    BorderColor =4210752
                    Name ="Line69"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =2370
                    Width =10545
                    BorderColor =4210752
                    Name ="Line32"
                End
                Begin Label
                    TextAlign =2
                    Left =2430
                    Top =2490
                    Width =1075
                    Height =227
                    FontSize =8
                    FontWeight =400
                    Name ="lblBalBF"
                    Caption ="Balance B/F"
                End
                Begin Line
                    Left =8400
                    Top =2430
                    Width =0
                    Height =340
                    Name ="Line41"
                End
                Begin Line
                    Left =1425
                    Top =2430
                    Width =0
                    Height =340
                    Name ="Line42"
                End
                Begin Line
                    Left =2370
                    Top =2430
                    Width =0
                    Height =340
                    Name ="Line43"
                End
                Begin Line
                    Left =4770
                    Top =2430
                    Width =0
                    Height =340
                    Name ="Line44"
                End
                Begin Line
                    Left =7215
                    Top =2430
                    Width =0
                    Height =340
                    Name ="Line45"
                End
                Begin Line
                    Top =2770
                    Width =10665
                    Name ="Line46"
                End
                Begin Line
                    Left =3514
                    Top =2430
                    Width =0
                    Height =340
                    Name ="Line47"
                End
                Begin Line
                    Left =9540
                    Top =2430
                    Width =0
                    Height =340
                    Name ="Line58"
                End
                Begin Label
                    TextAlign =3
                    Left =7215
                    Top =2490
                    Width =1150
                    Height =227
                    FontSize =8
                    FontWeight =400
                    Name ="lblBalanceBFDR"
                End
                Begin Label
                    TextAlign =3
                    Left =8400
                    Top =2490
                    Width =1135
                    Height =227
                    FontSize =8
                    FontWeight =400
                    Name ="lblBalanceBFCR"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =855
                    Top =623
                    Width =1134
                    TabIndex =4
                    Name ="Text72"
                    ControlSource ="=[Bal]+[DR]+[CR]+[OpenBalDR]-[OpenBalDR]"
                    Format ="Standard"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2085
                    Top =1260
                    Height =375
                    FontSize =12
                    FontWeight =700
                    TabIndex =5
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
                    TabIndex =6
                    Name ="accName"
                    ControlSource ="accName"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1410
                    Top =2085
                    Width =956
                    Height =227
                    FontSize =10
                    Name ="Label18"
                    Caption ="Doc Date"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =4980
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
                    Left =2475
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
                    Left =450
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
                    Left =3660
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
                    Left =7485
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
                    Left =8730
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
                    TabIndex =7
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
                    TabIndex =8
                    Name ="PeriodTo"
                    ControlSource ="=Forms!frmReport!txtToStr2 & ' (' & IIf(Forms!frmReport!txtToStr2=0,0,MonthName("
                        "Forms!frmReport!txtToStr2)) & ')'"
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
                    TabIndex =9
                    Name ="VatType"
                    ControlSource ="VatType"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =345
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1485
                    Top =60
                    Width =906
                    TabIndex =1
                    Name ="trnsDocDate"
                    ControlSource ="trnsDocDate"
                    Format ="Short Date"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2430
                    Top =60
                    Width =1071
                    Name ="trnsNote"
                    ControlSource ="refName"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3510
                    Top =60
                    Width =1251
                    TabIndex =2
                    Name ="docNo"
                    ControlSource ="docNo"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4935
                    Top =56
                    Width =2226
                    ColumnWidth =6315
                    TabIndex =3
                    Name ="desName"
                    ControlSource ="trnsNote"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7215
                    Top =60
                    Width =1134
                    TabIndex =4
                    Name ="txttrnsDebits"
                    ControlSource ="=IIf(IsNumeric([trnsDebits]),IIf([trnsDebits]=0,\"\",Abs([trnsDebits])),\"\")"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8400
                    Top =60
                    Width =1134
                    TabIndex =5
                    Name ="txttrnsCredits"
                    ControlSource ="=IIf(IsNumeric([trnsCredits]),IIf([trnsCredits]=0,\"\",[trnsCredits]),\"\")"
                    Format ="Standard"
                End
                Begin Line
                    Left =8400
                    Width =0
                    Height =340
                    Name ="Line11"
                End
                Begin Line
                    Left =1425
                    Width =0
                    Height =340
                    Name ="Line13"
                End
                Begin Line
                    Left =2385
                    Width =0
                    Height =340
                    Name ="Line14"
                End
                Begin Line
                    Left =4770
                    Width =0
                    Height =340
                    Name ="Line16"
                End
                Begin Line
                    Left =7218
                    Width =0
                    Height =340
                    Name ="Line17"
                End
                Begin Line
                    Top =345
                    Width =10665
                    Name ="Line27"
                End
                Begin Line
                    Left =3510
                    Width =0
                    Height =340
                    Name ="Line30"
                End
                Begin Line
                    Left =9540
                    Width =0
                    Height =340
                    Name ="Line59"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =1
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4770
                    Top =60
                    Width =1134
                    TabIndex =6
                    Name ="RunTotal"
                    ControlSource ="=([trnsDebits]-[trnsCredits])"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =60
                    Width =1134
                    TabIndex =7
                    Name ="RunBal"
                    ControlSource ="=[RunTotal]+IIf(IsNumeric([Bal]-([DR]-[CR])),([Bal]-([DR]-[CR])),([OpenBalDR]+[O"
                        "penBalCR]))"
                    Format ="Standard"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1371
                    TabIndex =8
                    Name ="trnInternalRef"
                    ControlSource ="trnInternalRef"
                End
                Begin Subform
                    OldBorderStyle =0
                    Top =283
                    Width =10710
                    Height =62
                    TabIndex =9
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
                    Left =7095
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
                    Left =8340
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
                    Left =6240
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
                    Left =7095
                    Top =345
                    Width =1134
                    TabIndex =2
                    Name ="BalanceDB"
                    ControlSource ="=IIf([Bal]>=0,[Bal])"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =345
                    Width =1134
                    TabIndex =3
                    Name ="BalanceCR"
                    ControlSource ="=IIf([Bal]<0,Abs([Bal]))"
                    Format ="Standard"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =5265
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
                    Left =345
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
Public RunBalance, BalanceBF As Double
Public OldAccNo As String
Public LastEntry As String
Private Sub Report_NoData(Cancel As Integer)

MsgBox ("The report is empty.")
Cancel = True

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
   'RunBalance = Format(CDbl(IIf(IsNumeric(lblBalSum.Caption), lblBalSum.Caption, 0)) + CDbl(IIf(IsNull(trnsDebits), 0, trnsDebits)) - CDbl(IIf(IsNull(trnsCredits), 0, trnsCredits)), "##,###,###,##0.00")
   LastEntry = Format(CDbl(IIf((IsNull(trnsCredits) Or (trnsCredits = 0)), IIf(IsNull(trnsDebits), 0, trnsDebits), -trnsCredits)), "##,###,###,##0.00")
   OldAccNo = accNo
   
End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
   If accNo <> OldAccNo Then
      lblBalanceBFDR.Caption = IIf(IsNumeric([Bal] - ([DR] - [CR])), IIf([Bal] - ([DR] - [CR]) > 0, Format([Bal] - ([DR] - [CR]), "##,###,###,##0.00"), ""), IIf(IsNumeric([OpenBalDR]), IIf([OpenBalDR] = 0, "", Format([OpenBalDR], "##,###,###,##0.00")), ""))
      lblBalanceBFCR.Caption = IIf(IsNumeric([Bal] - ([DR] - [CR])), IIf([Bal] - ([DR] - [CR]) < 0, Format(Abs([Bal] - ([DR] - [CR])), "##,###,###,##0.00"), ""), IIf(IsNumeric([OpenBalCR]), IIf([OpenBalCR] = 0, "", Format(Abs([OpenBalCR]), "##,###,###,##0.00")), ""))
      If IsNumeric([Bal] - ([DR] - [CR])) Then
         RunBalance = ([Bal] - ([DR] - [CR]))
      Else
         If ([Bal] - ([DR] - [CR])) < 0 Then
            If IsNumeric([OpenBalCR]) Then
                BalanceBF = [OpenBalCR]
            Else
                BalanceBF = 0
            End If
         Else
            If IsNumeric([OpenBalDR]) Then
                BalanceBF = [OpenBalDR]
            Else
                BalanceBF = 0
            End If
         End If
      End If
      If Forms.frmReport.txtFromStr2.Value = 0 Then
        lblBalBF.Caption = "Op. Balance"
      Else
        lblBalBF.Caption = "Balance B/F"
      End If
   Else
      If Me.Pages > 1 And Me.Page > 1 Then
            lblBalBF.Caption = "Prev. Page Bal."
        If RunBal >= 0 Then
            lblBalanceBFDR.Caption = Format(CStr(RunBal - LastEntry), "##,###,###,##0.00")
            lblBalanceBFCR.Caption = ""
        Else
            lblBalanceBFDR.Caption = ""
            lblBalanceBFCR.Caption = Format(CStr(RunBal - LastEntry), "##,###,###,##0.00")
        End If
      End If
   End If
   
End Sub

Private Sub Report_Open(Cancel As Integer)
If IsNull(Forms.frmReport.cboFromStr1) = True Or Forms.frmReport.cboFromStr1 = "" Then
    Forms.frmReport.cboFromStr1.Value = DLookup("Min(accNo)", "tblAccount", "accIsVat = True") '"0"
End If
If IsNull(Forms.frmReport.cboToStr1) = True Or Forms.frmReport.cboToStr1 = "" Then
    Forms.frmReport.cboToStr1.Value = DLookup("Max(accNo)", "tblAccount", "accIsVat = True") '"9999999999"
End If
If IsNull(Forms.frmReport.txtFromDate1) = True Or Forms.frmReport.txtFromDate1 = "" Then
    Forms.frmReport.txtFromDate1.Value = "01-01-1900"
End If
If IsNull(Forms.frmReport.txtToDate1) = True Or Forms.frmReport.txtToDate1 = "" Then
    Forms.frmReport.txtToDate1.Value = Format(Date, "dd-mm-yyyy")
End If
If IsNull(Forms.frmReport.txtFromStr2) = True Or Forms.frmReport.txtFromStr2 = "" Then
    Forms.frmReport.txtFromStr2.Value = 0
End If
If IsNull(Forms.frmReport.txtToStr2) = True Or Forms.frmReport.txtToStr2 = "" Then
    Forms.frmReport.txtToStr2.Value = 12
End If

If Forms.frmReport.txtFromStr2.Value = 0 Then
    lblBalBF.Caption = "Op. Balance"
Else
    lblBalBF.Caption = "Balance B/F"
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


'Select Case Forms.frmReport.grpOptions
'    Case 1
'        Me.Filter = "Bal = 0"
'        Me.FilterOn = True
'    Case 2
'        Me.Filter = "CR Is Null And DR Is Null"
'       Me.FilterOn = True
'    Case 3
'        Me.Filter = "CR >0 Or DR > 0"
'        Me.FilterOn = True
'    Case 4
'        Me.Filter = ""
'        Me.FilterOn = True
'End Select

Select Case Forms.frmReport.grpOptions
    Case 1
        Me.Filter = "(Bal = 0) OR (IsNumeric(Bal)=False) "
        Me.FilterOn = True
    Case 2
        Me.Filter = "CR Is Null And DR Is Null"
        Me.FilterOn = True
    Case 3
        Me.Filter = "(CR >0 Or DR > 0)"
        Me.FilterOn = True
    Case 4
        
        Me.Filter = "(Bal <> 0 OR openbalDR <> 0 or openbalCR <> 0)"    '&& added 20/03/15 openbalDR <> 0 or openbalCR <> 0
                                                                        '&& to show the ones with only opening balance

        'Me.Filter = "Bal <> 0"
        Me.FilterOn = True
    Case Else
        Me.Filter = ""
        Me.FilterOn = False
End Select
 


End Sub
