Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10827
    DatasheetFontHeight =10
    ItemSuffix =34
    Left =2085
    Top =135
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    OrderBy ="tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChart"
        ".coaRef5"
    RecSrcDt = Begin
        0xa67c439bbc7ae340
    End
    RecordSource ="qryTrialBalance3"
    Caption ="Trial Balance"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x27010000370200003702000037020000000000004b2a00000801000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
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
            BorderLineStyle =0
            Width =850
            Height =850
            BorderColor =12632256
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            BorderLineStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontCharSet =238
            TextFontFamily =18
            BorderLineStyle =0
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
            BorderLineStyle =0
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
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman CE"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1417
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =161
                    TextAlign =2
                    TextFontFamily =18
                    Left =1980
                    Top =420
                    Width =5925
                    Height =645
                    FontSize =26
                    Name ="Label10"
                    Caption ="Trial Balance"
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
                    Left =8822
                    Top =731
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
                    Left =8565
                    Top =450
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
                    Left =10329
                    Top =450
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
                    Width =10827
                    BorderColor =4210752
                    Name ="Line69"
                End
            End
        End
        Begin PageHeader
            Height =793
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1485
                    Top =225
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
                    Top =225
                    Width =1425
                    Height =270
                    Name ="coaRef_Label"
                    Caption ="Reference"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =4530
                    Top =225
                    Width =4432
                    Height =270
                    Name ="accName_Label"
                    Caption ="Account"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =735
                    Width =10827
                    BorderColor =4210752
                    Name ="Line25"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =9930
                    Top =225
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
                    Left =9015
                    Top =225
                    Width =885
                    Height =270
                    Name ="Label29"
                    Caption ="Credits"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =264
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1485
                    Width =3000
                    Height =260
                    ColumnWidth =6315
                    FontSize =8
                    Name ="coaName"
                    ControlSource ="coaName"
                    FontName ="Arial"

                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Top =4
                    Width =1426
                    Height =260
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
                    Left =4530
                    Width =675
                    Height =260
                    FontSize =8
                    TabIndex =2
                    Name ="accNo"
                    ControlSource ="accNo"
                    FontName ="Arial"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5265
                    Width =3682
                    Height =260
                    FontSize =8
                    TabIndex =3
                    Name ="accName"
                    ControlSource ="accName"
                    FontName ="Arial"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9030
                    Width =851
                    Height =260
                    FontSize =8
                    TabIndex =4
                    Name ="Text26"
                    ControlSource ="=IIf((IIf(IsNumeric([CR]),[CR]+[OpenBalCR],[OpenBalCR])-IIf(IsNumeric([DR]),[DR]"
                        "-[OpenBalDR],-[OpenBalDR]))>0,(IIf(IsNumeric([CR]),[CR]+[OpenBalCR],[OpenBalCR])"
                        "-IIf(IsNumeric([DR]),[DR]-[OpenBalDR],-[OpenBalDR])))"
                    Format ="Standard"
                    FontName ="Arial"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9945
                    Width =851
                    Height =260
                    FontSize =8
                    TabIndex =5
                    Name ="txtDR"
                    ControlSource ="=IIf((IIf(IsNumeric([CR]),[CR]+[OpenBalCR],[OpenBalCR])-IIf(IsNumeric([DR]),[DR]"
                        "-[OpenBalDR],-[OpenBalDR]))<0,(IIf(IsNumeric([CR]),[CR]+[OpenBalCR],[OpenBalCR])"
                        "-IIf(IsNumeric([DR]),[DR]-[OpenBalDR],-[OpenBalDR])))"
                    Format ="Standard"
                    FontName ="Arial"

                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =650
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9930
                    Width =851
                    Height =260
                    FontSize =8
                    Name ="Text31"
                    ControlSource ="=Sum([DR])-Sum([OpenBalDR])"
                    Format ="Standard"
                    FontName ="Arial"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9015
                    Width =851
                    Height =260
                    FontSize =8
                    TabIndex =1
                    Name ="Text33"
                    ControlSource ="=Sum([CR])+Sum([OpenBalCR])"
                    Format ="Standard"
                    FontName ="Arial"

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
        Me.FilterOn = False
End Select
    


End Sub
