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
    Width =10341
    DatasheetFontHeight =10
    ItemSuffix =37
    Left =2520
    Top =180
    DatasheetGridlinesColor =12632256
    Toolbar ="PrintReportEmail"
    RecSrcDt = Begin
        0x2ba1a7213b8ae340
    End
    RecordSource ="qryStatement"
    Caption ="Statement Report"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x530300003702000037020000370200000000000065280000f400000001000000 ,
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
        Begin PageHeader
            Height =2154
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1695
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
                    Left =3508
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
                    Left =855
                    Top =1755
                    Width =1016
                    Height =227
                    FontSize =10
                    Name ="Label18"
                    Caption ="Doc. Date"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =4485
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
                    Left =1980
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
                    Top =1755
                    Width =795
                    Height =240
                    FontSize =10
                    Name ="Label21"
                    Caption ="Vaucher"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =3225
                    Top =1755
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
                    Left =7200
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
                    Left =8670
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
                    Left =1995
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
                    Left =2730
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
                    Left =8162
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
                    Left =7905
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
                    Left =9669
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
                    BorderWidth =2
                    Top =1174
                    Width =10260
                    BorderColor =4210752
                    Name ="Line69"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =18
                    Left =60
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
                    Width =10260
                    BorderColor =4210752
                    Name ="Line32"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =244
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =915
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
                    Left =1980
                    Top =4
                    Width =1071
                    Name ="trnsNote"
                    ControlSource ="refName"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3180
                    Top =4
                    Width =1251
                    TabIndex =2
                    Name ="refName"
                    ControlSource ="desName"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =60
                    Top =4
                    Width =561
                    TabIndex =3
                    Name ="trnInternalRef"
                    ControlSource ="trnInternalRef"
                End
                Begin TextBox
                    OverlapFlags =12
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4527
                    Width =2511
                    ColumnWidth =6315
                    TabIndex =4
                    Name ="desName"
                    ControlSource ="trnsNote"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7155
                    Top =4
                    Width =1416
                    TabIndex =5
                    Name ="trnsDebits"
                    ControlSource ="trnsDebits"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8670
                    Top =4
                    Width =1566
                    TabIndex =6
                    Name ="trnsCredits"
                    ControlSource ="trnsCredits"
                End
                Begin Line
                    Left =8625
                    Width =0
                    Height =240
                    Name ="Line11"
                End
                Begin Line
                    Left =855
                    Width =0
                    Height =240
                    Name ="Line13"
                End
                Begin Line
                    Left =1920
                    Width =0
                    Height =240
                    Name ="Line14"
                End
                Begin Line
                    Left =4485
                    Width =0
                    Height =240
                    Name ="Line16"
                End
                Begin Line
                    Left =7095
                    Width =0
                    Height =240
                    Name ="Line17"
                End
                Begin Line
                    Left =60
                    Top =240
                    Width =10260
                    Name ="Line27"
                End
                Begin Line
                    Left =3120
                    Width =0
                    Height =240
                    Name ="Line30"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =453
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7155
                    Top =60
                    Width =1416
                    Name ="TotalDR"
                    ControlSource ="=Sum([trnsDebits])"
                    Format ="#,##0.00\" zl\";-#,##0.00\" zl\""
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8670
                    Top =60
                    Width =1566
                    TabIndex =1
                    Name ="TotalCR"
                    ControlSource ="=Sum([trnsCredits])"
                    Format ="#,##0.00\" zl\";-#,##0.00\" zl\""
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =6300
                    Top =60
                    Width =737
                    Height =227
                    FontSize =10
                    Name ="Label36"
                    Caption ="Totals:"
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

Private Sub Report_Open(Cancel As Integer)
If IsNull(Forms.frmReport.txtFromStr1) = True Or Forms.frmReport.txtFromStr1 = "" Then
    Forms.frmReport.txtFromStr1.Value = "0"
End If
If IsNull(Forms.frmReport.txtToStr1) = True Or Forms.frmReport.txtToStr1 = "" Then
    Forms.frmReport.txtToStr1.Value = "9999999999"
End If
If IsNull(Forms.frmReport.txtFromDate1) = True Or Forms.frmReport.txtFromDate1 = "" Then
    Forms.frmReport.txtFromDate1.Value = "01-01-1900"
End If
If IsNull(Forms.frmReport.txtToDate1) = True Or Forms.frmReport.txtToDate1 = "" Then
    Forms.frmReport.txtToDate1.Value = Format(Date, "dd-mm-yyyy")
End If

End Sub
