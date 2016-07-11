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
    Width =10719
    DatasheetFontHeight =10
    ItemSuffix =74
    Left =345
    Top =105
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    Filter ="(CR >0 Or DR > 0)"
    Toolbar ="PrintReportEmail"
    RecSrcDt = Begin
        0xf9a202adef8ae340
    End
    RecordSource ="qryStatementForBal"
    Caption ="Statement Report"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x3702000037020000370200003702000000000000df290000f400000001000000 ,
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
            Height =2485
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1755
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
                    Left =3568
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
                    Top =1755
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
                    Left =4755
                    Top =1755
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
                    Top =1755
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
                    Left =390
                    Top =1755
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
                    Left =3570
                    Top =1755
                    Width =765
                    Height =270
                    FontSize =10
                    Name ="Label22"
                    Caption ="Doc. No"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =7365
                    Top =1755
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
                    Top =1755
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
                    Left =2385
                    Top =420
                    Width =5895
                    Height =645
                    FontSize =26
                    Name ="Label10"
                    Caption ="Statement"
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
                    Height =30
                    BorderColor =4210752
                    Name ="Line69"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =18
                    Left =120
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
                    Top =2085
                    Width =10650
                    BorderColor =4210752
                    Name ="Line32"
                End
                Begin Label
                    TextAlign =2
                    Left =2205
                    Top =2205
                    Width =1180
                    Height =227
                    FontSize =8
                    FontWeight =400
                    Name ="lblBalBF"
                    Caption ="Balance B/F"
                End
                Begin Line
                    Left =8325
                    Top =2145
                    Width =0
                    Height =340
                    Name ="Line41"
                End
                Begin Line
                    Left =1129
                    Top =2145
                    Width =0
                    Height =340
                    Name ="Line42"
                End
                Begin Line
                    Left =2206
                    Top =2145
                    Width =0
                    Height =340
                    Name ="Line43"
                End
                Begin Line
                    Left =4365
                    Top =2145
                    Width =0
                    Height =340
                    Name ="Line44"
                End
                Begin Line
                    Left =7081
                    Top =2145
                    Width =0
                    Height =340
                    Name ="Line45"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =2485
                    Width =10710
                    Name ="Line46"
                End
                Begin Line
                    Left =3394
                    Top =2145
                    Width =0
                    Height =340
                    Name ="Line47"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =8205
                    Top =1200
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
                    Top =1425
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
                    Left =9465
                    Top =1200
                    Width =1086
                    TabIndex =6
                    Name ="PeriodFrom"
                    ControlSource ="=Forms!frmReport!txtFromStr2 & ' (' & IIf(Forms!frmReport!txtFromStr2=0,0,MonthN"
                        "ame(Forms!frmReport!txtFromStr2)) & ')'"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9465
                    Top =1425
                    Width =1086
                    TabIndex =7
                    Name ="PeriodTo"
                    ControlSource ="=Forms!frmReport!txtToStr2 & ' (' & IIf(Forms!frmReport!txtToStr2=0,0,MonthName("
                        "Forms!frmReport!txtToStr2)) & ')'"
                End
                Begin Line
                    Left =9525
                    Top =2145
                    Width =0
                    Height =340
                    Name ="Line58"
                End
                Begin Label
                    TextAlign =3
                    Left =7140
                    Top =2205
                    Width =1150
                    Height =227
                    FontSize =8
                    FontWeight =400
                    Name ="lblBalanceBFDR"
                End
                Begin Label
                    TextAlign =3
                    Left =8385
                    Top =2205
                    Width =1135
                    Height =227
                    FontSize =8
                    FontWeight =400
                    Name ="lblBalanceBFCR"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =9690
                    Top =1755
                    Width =765
                    Height =270
                    FontSize =10
                    Name ="Label66"
                    Caption ="Balance"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1693
                    Top =56
                    Width =1134
                    TabIndex =8
                    Name ="Text72"
                    ControlSource ="=[Bal]+[DR]+[CR]+[OpenBalDR]-[OpenBalDR]"
                    Format ="Standard"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =390
                    Top =690
                    Height =375
                    FontSize =12
                    FontWeight =700
                    TabIndex =9
                    Name ="crnShorName"
                    ControlSource ="=DLookUp(\"crnShortName\",\"tblCurrency\",\"[crnID] = \" & [accCur])"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =244
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1185
                    Top =4
                    Width =906
                    TabIndex =1
                    Name ="trnsDocDate"
                    ControlSource ="trnsDocDate"
                    Format ="Short Date"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2250
                    Top =4
                    Width =1071
                    Name ="trnsNote"
                    ControlSource ="refName"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3390
                    Top =4
                    Width =920
                    TabIndex =2
                    Name ="docNo"
                    ControlSource ="docNo"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Top =4
                    Width =1146
                    TabIndex =3
                    Name ="trnInternalRef"
                    ControlSource ="trnInternalRef"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4425
                    Width =2631
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
                    Top =4
                    Width =1134
                    TabIndex =5
                    Name ="txttrnsDebits"
                    ControlSource ="=IIf(IsNumeric([trnsDebits]),IIf([trnsDebits]=0,\"\",Abs([trnsDebits])),\"\")"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8385
                    Top =4
                    Width =1134
                    TabIndex =6
                    Name ="txttrnsCredits"
                    ControlSource ="=IIf(IsNumeric([trnsCredits]),IIf([trnsCredits]=0,\"\",[trnsCredits]),\"\")"
                    Format ="Standard"
                End
                Begin Line
                    Left =8325
                    Width =0
                    Height =240
                    Name ="Line11"
                End
                Begin Line
                    Left =1125
                    Width =0
                    Height =240
                    Name ="Line13"
                End
                Begin Line
                    Left =2205
                    Width =0
                    Height =240
                    Name ="Line14"
                End
                Begin Line
                    Left =4365
                    Width =0
                    Height =240
                    Name ="Line16"
                End
                Begin Line
                    Left =7080
                    Width =0
                    Height =240
                    Name ="Line17"
                End
                Begin Line
                    Left =-4
                    Top =240
                    Width =10710
                    Name ="Line27"
                End
                Begin Line
                    Left =3390
                    Width =0
                    Height =240
                    Name ="Line30"
                End
                Begin Line
                    Left =9525
                    Width =0
                    Height =240
                    Name ="Line59"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =1
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5670
                    Width =1134
                    TabIndex =7
                    Name ="RunTotal"
                    ControlSource ="=([trnsDebits]-[trnsCredits])"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9585
                    Top =4
                    Width =1134
                    TabIndex =8
                    Name ="RunBal"
                    ControlSource ="=[RunTotal]+IIf(IsNumeric([Bal]-([DR]-[CR])),([Bal]-([DR]-[CR])),([OpenBalDR]+[O"
                        "penBalCR]))"
                    Format ="Standard"
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
                    ControlSource ="=IIf([Bal]>=0,[Bal])"
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
                    ControlSource ="=IIf([Bal]<0,Abs([Bal]))"
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
Public RunBalance, BalanceBF As Double
Public OldAccNo As String
Public LastEntry As String
Private Sub Report_NoData(Cancel As Integer)

MsgBox ("The report is empty.")
Cancel = True

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
   'RunBalance = Format(CDbl(IIf(IsNumeric(lblBalSum.Caption), lblBalSum.Caption, 0)) + CDbl(IIf(IsNull(trnsDebits), 0, trnsDebits)) - CDbl(IIf(IsNull(trnsCredits), 0, trnsCredits)), "##########0.00")
   LastEntry = Format(CDbl(IIf((IsNull(trnsCredits) Or (trnsCredits = 0)), IIf(IsNull(trnsDebits), 0, trnsDebits), -trnsCredits)), "##########0.00")
   OldAccNo = accNo
   
End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
   If accNo <> OldAccNo Then
      lblBalanceBFDR.Caption = IIf(IsNumeric([Bal] - ([DR] - [CR])), IIf([Bal] - ([DR] - [CR]) > 0, Format([Bal] - ([DR] - [CR]), "##########0.00"), ""), IIf(IsNumeric([OpenBalDR]), IIf([OpenBalDR] = 0, "", Format([OpenBalDR], "##########0.00")), ""))
      lblBalanceBFCR.Caption = IIf(IsNumeric([Bal] - ([DR] - [CR])), IIf([Bal] - ([DR] - [CR]) < 0, Format(Abs([Bal] - ([DR] - [CR])), "##########0.00"), ""), IIf(IsNumeric([OpenBalCR]), IIf([OpenBalCR] = 0, "", Format(Abs([OpenBalCR]), "##########0.00")), ""))
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
        If RunBal >= 0 Then
            lblBalanceBFDR.Caption = CStr(RunBal - LastEntry)
            lblBalanceBFCR.Caption = ""
            lblBalBF.Caption = "Prev. Page Bal."
        Else
            lblBalanceBFDR.Caption = ""
            lblBalanceBFCR.Caption = CStr(RunBal - LastEntry)
            lblBalBF.Caption = "Prev. Page Bal."
        End If
      End If
   End If
   
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
        Me.Filter = "Bal <> 0"
        Me.FilterOn = True
    Case Else
        Me.Filter = ""
        Me.FilterOn = False
End Select
 


End Sub
