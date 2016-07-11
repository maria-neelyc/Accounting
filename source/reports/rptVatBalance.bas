Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10318
    DatasheetFontHeight =10
    ItemSuffix =41
    Left =2148
    Top =228
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    OrderBy ="tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChart"
        ".coaRef5"
    Toolbar ="PrintReportEmail"
    RecSrcDt = Begin
        0xb8f03387799ae340
    End
    RecordSource ="qryVatBal3"
    Caption ="VAT"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x53030000370200003702000037020000000000004e2800008b02000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =238
            TextAlign =1
            FontSize =10
            FontWeight =700
            FontName ="Arial CE"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =2
            Width =850
            Height =850
            BorderColor =12632256
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
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontCharSet =238
            TextFontFamily =18
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman CE"
            AsianLineBreak =255
        End
        Begin ListBox
            TextFontCharSet =238
            TextFontFamily =18
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman CE"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =238
            TextFontFamily =18
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman CE"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =2
            ControlSource ="VatType"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1695
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =161
                    TextAlign =2
                    TextFontFamily =18
                    Left =1965
                    Top =420
                    Width =5955
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
                    Left =2730
                    Width =4545
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
                    Left =8162
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
                    Left =7905
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
                    Left =9669
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
                    BorderWidth =2
                    Top =1174
                    Width =10260
                    BorderColor =4210752
                    Name ="Line69"
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1200
            Name ="GroupHeader2"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1485
                    Top =690
                    Width =3000
                    Height =270
                    Name ="coaName_Label"
                    Caption ="Ref. Heading"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Top =690
                    Width =1425
                    Height =270
                    Name ="coaRef_Label"
                    Caption ="Reference"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =9420
                    Top =690
                    Width =851
                    Height =270
                    Name ="Bal_Label"
                    Caption ="Balance"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =4530
                    Top =690
                    Width =3022
                    Height =270
                    Name ="accName_Label"
                    Caption ="Account"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =1200
                    Width =10260
                    BorderColor =4210752
                    Name ="Line25"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =8490
                    Top =690
                    Width =885
                    Height =270
                    Name ="Label28"
                    Caption ="Debits"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =7575
                    Top =690
                    Width =885
                    Height =270
                    Name ="Label29"
                    Caption ="Credits"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =850
                    Top =165
                    Width =1426
                    Height =271
                    FontSize =10
                    FontWeight =700
                    Name ="VatType"
                    ControlSource ="VatType"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontFamily =18
                            Left =120
                            Top =165
                            Width =570
                            Height =270
                            Name ="Label30"
                            Caption ="VAT"
                            FontName ="Times New Roman"
                            Tag ="DetachedLabel"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =651
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1485
                    Width =3000
                    Height =315
                    ColumnWidth =6315
                    FontSize =8
                    Name ="coaName"
                    ControlSource ="coaName"
                    FontName ="Arial"
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =1426
                    Height =315
                    FontSize =8
                    TabIndex =1
                    Name ="coaRef"
                    ControlSource ="coaRef"
                    FontName ="Arial"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9420
                    Width =851
                    Height =315
                    FontSize =8
                    TabIndex =2
                    Name ="Bal"
                    ControlSource ="Bal"
                    Format ="Standard"
                    FontName ="Arial"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4530
                    Width =675
                    Height =315
                    FontSize =8
                    TabIndex =3
                    Name ="accNo"
                    ControlSource ="accNo"
                    FontName ="Arial"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5265
                    Width =2272
                    Height =315
                    FontSize =8
                    TabIndex =4
                    Name ="accName"
                    ControlSource ="accName"
                    FontName ="Arial"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7590
                    Width =851
                    Height =315
                    FontSize =8
                    TabIndex =5
                    Name ="Text26"
                    ControlSource ="CR"
                    Format ="Standard"
                    FontName ="Arial"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8505
                    Width =851
                    Height =315
                    FontSize =8
                    TabIndex =6
                    Name ="Text27"
                    ControlSource ="DR"
                    Format ="Standard"
                    FontName ="Arial"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7596
                    Top =336
                    Width =851
                    Height =315
                    FontSize =8
                    TabIndex =7
                    Name ="V_CR"
                    ControlSource ="V_CR"
                    Format ="Standard"
                    FontName ="Arial"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8511
                    Top =336
                    Width =851
                    Height =315
                    FontSize =8
                    TabIndex =8
                    Name ="V_DR"
                    ControlSource ="V_DR"
                    Format ="Standard"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =4
                    TextAlign =3
                    Left =6240
                    Top =336
                    Width =1332
                    Height =252
                    FontWeight =400
                    Name ="Label40"
                    Caption ="Vatable Value:"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =623
            Name ="GroupFooter0"
            Begin
                Begin Line
                    Left =60
                    Top =390
                    Width =10205
                    Name ="Line33"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9420
                    Top =60
                    Width =851
                    Height =315
                    FontSize =8
                    Name ="Total"
                    ControlSource ="=Sum([Bal])"
                    Format ="Standard"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =2
                            TextFontFamily =18
                            Left =8505
                            Top =60
                            Width =851
                            Height =270
                            Name ="Label37"
                            Caption ="Total:"
                            FontName ="Times New Roman"
                            Tag ="DetachedLabel"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =510
            Name ="ReportFooter"
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
If IsNull(Forms.frmReport.txtFromStr1) = True Or Forms.frmReport.txtFromStr1 = "" Then
    Forms.frmReport.txtFromStr1.Value = 1
End If
If IsNull(Forms.frmReport.txtToStr1) = True Or Forms.frmReport.txtToStr1 = "" Then
    Forms.frmReport.txtToStr1.Value = 12
End If
If IsNull(Forms.frmReport.txtToStr2) = True Or Forms.frmReport.txtToStr2 = "" Then
    Forms.frmReport.txtToStr2.Value = 5
End If

If Forms.frmReport.cboOrderBy = 1 Then
    Me.OrderBy = "tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChart.coaRef5"
Else
    Me.OrderBy = "qryTriaBalace2.accNo"
End If

Select Case Forms.frmReport.grpOptions
    Case 1
    Case 2
        DoCmd.OpenReport "rptVatStatement", acViewPreview
    Case 3
    Case 4
End Select
    


End Sub
