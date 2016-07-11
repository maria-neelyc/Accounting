Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15649
    DatasheetFontHeight =10
    ItemSuffix =105
    Left =1005
    Top =495
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    Toolbar ="PrintReportEmail"
    RecSrcDt = Begin
        0x7461dfd8568ce340
    End
    RecordSource ="qryTrnsPages"
    Caption ="Statement Report"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x54010000530300004d0300003702000000000000213d0000e301000001000000 ,
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
            ControlSource ="trnPageCounter"
        End
        Begin BreakLevel
            ControlSource ="trnInternalRef"
        End
        Begin BreakLevel
            ControlSource ="persID"
        End
        Begin BreakLevel
            ControlSource ="trnEntryDate"
        End
        Begin PageHeader
            Height =2955
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1200
                    Top =1260
                    Height =375
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="trnPageCounter"
                    ControlSource ="trnPageCounter"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4140
                    Top =1260
                    Width =2451
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="trnInternalRef"
                    ControlSource ="trnInternalRef"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontCharSet =0
                            TextFontFamily =18
                            Left =3000
                            Top =1260
                            Width =1020
                            Height =300
                            FontSize =12
                            Name ="Label73"
                            Caption ="Voucher:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =566
                    Top =2387
                    Width =1080
                    Height =525
                    FontSize =10
                    TopMargin =113
                    Name ="Label21"
                    Caption ="Account "
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =161
                    TextAlign =2
                    TextFontFamily =18
                    Left =4920
                    Top =390
                    Width =5925
                    Height =645
                    FontSize =26
                    Name ="Label10"
                    Caption ="Transactions Per Page"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextFontCharSet =161
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =60
                    Width =15405
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
                    Left =13517
                    Top =731
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
                    Left =13260
                    Top =450
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
                    Left =15024
                    Top =450
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
                    BorderWidth =2
                    Top =1174
                    Width =15647
                    BorderColor =4210752
                    Name ="Line69"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =18
                    Left =60
                    Top =1260
                    Width =1020
                    Height =300
                    FontSize =12
                    Name ="Label63"
                    Caption ="Page No:"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =13155
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
                    Left =13155
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
                    Left =14415
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
                    Left =14415
                    Top =1425
                    Width =1086
                    TabIndex =7
                    Name ="PeriodTo"
                    ControlSource ="=Forms!frmReport!txtToStr2 & ' (' & IIf(Forms!frmReport!txtToStr2=0,0,MonthName("
                        "Forms!frmReport!txtToStr2)) & ')'"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7657
                    Top =1273
                    Width =1266
                    Height =375
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    Name ="persID"
                    ControlSource ="persID"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontCharSet =0
                            TextFontFamily =18
                            Left =6690
                            Top =1260
                            Width =915
                            Height =300
                            FontSize =12
                            Name ="Label75"
                            Caption ="Period:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10660
                    Top =1260
                    Width =1761
                    Height =375
                    FontSize =12
                    FontWeight =700
                    TabIndex =9
                    Name ="trnEntryDate"
                    ControlSource ="trnEntryDate"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontCharSet =0
                            TextFontFamily =18
                            Left =9135
                            Top =1262
                            Width =1350
                            Height =300
                            FontSize =12
                            Name ="Label77"
                            Caption ="Entry Date:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =3525
                    Top =2385
                    Width =855
                    Height =525
                    FontSize =10
                    Name ="Label90"
                    Caption ="Transact.Date"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =4365
                    Top =2385
                    Width =555
                    Height =525
                    FontSize =10
                    TopMargin =113
                    Name ="Label91"
                    Caption ="VAT"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =4920
                    Top =2385
                    Width =555
                    Height =525
                    FontSize =10
                    Name ="Label92"
                    Caption ="Reference"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =5490
                    Top =2385
                    Width =1200
                    Height =525
                    FontSize =10
                    Name ="Label93"
                    Caption ="Document No."
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =6690
                    Top =2385
                    Width =1005
                    Height =525
                    FontSize =10
                    Name ="Label94"
                    Caption ="Transact. Date"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =7755
                    Top =2385
                    Width =2730
                    Height =525
                    FontSize =10
                    TopMargin =113
                    Name ="Label95"
                    Caption ="Notes"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =10545
                    Top =2385
                    Width =1140
                    Height =525
                    FontSize =10
                    TopMargin =113
                    Name ="Label96"
                    Caption ="Debits"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =11685
                    Top =2385
                    Width =1140
                    Height =525
                    FontSize =10
                    TopMargin =113
                    Name ="Label97"
                    Caption ="Credits"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =12855
                    Top =2385
                    Width =465
                    Height =525
                    FontSize =10
                    TopMargin =113
                    Name ="Label98"
                    Caption ="Curr"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =13320
                    Top =2385
                    Width =1140
                    Height =525
                    FontSize =10
                    Name ="Label99"
                    Caption ="Foreign Debits"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =14460
                    Top =2385
                    Width =1140
                    Height =525
                    FontSize =10
                    Name ="Label100"
                    Caption ="Foreign Credits"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =2
                    Top =2955
                    Width =15648
                    BorderColor =4210752
                    Name ="Line32"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =483
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =915
                    Top =60
                    Width =2601
                    Height =414
                    TabIndex =1
                    Name ="trnsDate"
                    ControlSource ="accName"
                    Format ="Short Date"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4920
                    Top =60
                    Width =576
                    Height =414
                    Name ="refShortName"
                    ControlSource ="refShortName"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5490
                    Top =60
                    Width =1251
                    Height =414
                    TabIndex =2
                    Name ="docNo"
                    ControlSource ="docNo"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Top =60
                    Width =921
                    Height =414
                    TabIndex =3
                    Name ="accNo"
                    ControlSource ="accNo"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7755
                    Top =60
                    Width =2736
                    Height =414
                    ColumnWidth =6315
                    TabIndex =4
                    Name ="desName"
                    ControlSource ="trnsNote"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10549
                    Top =56
                    Width =1134
                    Height =414
                    TabIndex =5
                    Name ="txttrnsDebits"
                    ControlSource ="=IIf(IsNumeric([trnsDebits]),IIf([trnsDebits]=0,0,Abs([trnsDebits])),0)"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11683
                    Top =56
                    Width =1134
                    Height =414
                    TabIndex =6
                    Name ="txttrnsCredits"
                    ControlSource ="=IIf(IsNumeric([trnsCredits]),IIf([trnsCredits]=0,0,[trnsCredits]),0)"
                    Format ="Standard"
                End
                Begin Line
                    Left =6735
                    Width =0
                    Height =483
                    Name ="Line11"
                End
                Begin Line
                    Left =915
                    Width =0
                    Height =483
                    Name ="Line13"
                End
                Begin Line
                    Left =3510
                    Width =0
                    Height =483
                    Name ="Line14"
                End
                Begin Line
                    Left =4920
                    Width =0
                    Height =483
                    Name ="Line16"
                End
                Begin Line
                    Left =5490
                    Width =0
                    Height =483
                    Name ="Line17"
                End
                Begin Line
                    Top =483
                    Width =15593
                    Name ="Line27"
                End
                Begin Line
                    Left =4425
                    Width =0
                    Height =483
                    Name ="Line30"
                End
                Begin Line
                    Left =7710
                    Width =0
                    Height =483
                    Name ="Line59"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3510
                    Top =60
                    Width =921
                    Height =414
                    TabIndex =7
                    Name ="Text78"
                    ControlSource ="trnsDate"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4425
                    Top =60
                    Width =501
                    Height =414
                    TabIndex =8
                    Name ="vatName"
                    ControlSource ="vatName"
                End
                Begin Line
                    Left =12810
                    Width =0
                    Height =483
                    Name ="Line80"
                End
                Begin Line
                    Left =10545
                    Width =0
                    Height =483
                    Name ="Line81"
                End
                Begin Line
                    Left =11685
                    Width =0
                    Height =483
                    Name ="Line82"
                End
                Begin Line
                    Left =13320
                    Width =0
                    Height =483
                    Name ="Line83"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6750
                    Top =60
                    Width =921
                    Height =414
                    TabIndex =9
                    Name ="trnsDocDate"
                    ControlSource ="trnsDocDate"
                End
                Begin Line
                    Left =14460
                    Width =0
                    Height =483
                    Name ="Line86"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12810
                    Top =60
                    Width =516
                    Height =414
                    TabIndex =10
                    Name ="crnShortName"
                    ControlSource ="crnShortName"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13320
                    Top =60
                    Width =1134
                    Height =414
                    TabIndex =11
                    Name ="FDebits"
                    ControlSource ="FDebits"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14454
                    Top =60
                    Width =1134
                    Height =414
                    TabIndex =12
                    Name ="FCredits"
                    ControlSource ="FCredits"
                    Format ="Standard"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =283
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10545
                    Width =1134
                    Height =279
                    Name ="TotalDR"
                    ControlSource ="=Sum([trnsDebits])"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11685
                    Width =1134
                    Height =279
                    TabIndex =1
                    Name ="TotalCR"
                    ControlSource ="=Sum([trnsCredits])"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13320
                    Width =1134
                    Height =279
                    TabIndex =2
                    Name ="TotalFDR"
                    ControlSource ="=Sum([FDebits])"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14460
                    Width =1134
                    Height =279
                    TabIndex =3
                    Name ="TotalFCR"
                    ControlSource ="=Sum([FCredits])"
                    Format ="Standard"
                End
            End
        End
        Begin PageFooter
            Height =170
            Name ="PageFooterSection"
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

Private Sub Report_Open(Cancel As Integer)

If FnIsLoad("frmTransaction") Then
    Me.RecordSource = "qryTrnsPagesOne"
    
    lblPrdTo.Visible = False
    lblPrdFrom.Visible = False
    PeriodFrom.Visible = False
    PeriodTo.Visible = False
Else
    Me.RecordSource = "qryTrnsPages"
    If IsNull(Forms.frmReport.txtFromStr1) = True Or Forms.frmReport.txtFromStr1 = "" Then
        Forms.frmReport.txtFromStr1.Value = "0"
    End If
    If IsNull(Forms.frmReport.txtToStr1) = True Or Forms.frmReport.txtToStr1 = "" Then
        Forms.frmReport.txtToStr1.Value = "9999999999"
    End If

    If IsNull(Forms.frmReport.txtFromStr2) = True Or Forms.frmReport.txtFromStr2 = "" Then
        Forms.frmReport.txtFromStr2.Value = 0
    End If
    If IsNull(Forms.frmReport.txtToStr2) = True Or Forms.frmReport.txtToStr2 = "" Then
        Forms.frmReport.txtToStr2.Value = 12
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


End If


'Select Case Forms.frmReport.grpOptions
'    Case 1
'        Me.Filter = "(Bal = 0) OR (IsNumeric(Bal)=False) "
'        Me.FilterOn = True
'    Case 2
'        Me.Filter = "CR Is Null And DR Is Null"
'        Me.FilterOn = True
'    Case 3
'        Me.Filter = "(CR >0 Or DR > 0)"
'        Me.FilterOn = True
'    Case 4
'        Me.Filter = "Bal <> 0"
'        Me.FilterOn = True
'    Case Else
'        Me.Filter = ""
'        Me.FilterOn = False
'End Select



End Sub
