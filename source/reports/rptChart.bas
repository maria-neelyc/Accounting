Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =10
    ItemSuffix =75
    Left =945
    Top =240
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    OrderBy ="coaRef"
    RecSrcDt = Begin
        0xc380d6e25578e340
    End
    RecordSource ="qryChartOfAccounts"
    Caption ="Chart of Accounts"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6e04000037020000370200003702000000000000602700004a01000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextFontCharSet =238
            TextAlign =1
            TextFontFamily =18
            FontSize =10
            FontWeight =700
            FontName ="Times New Roman"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderColor =8388608
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Line
            BorderColor =8388608
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            FontName ="Arial"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Subform
            OldBorderStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =2070
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =90
                    Top =1680
                    Width =1650
                    Height =300
                    Name ="Label68"
                    Caption ="Reference"
                End
                Begin Label
                    Left =1815
                    Top =1680
                    Width =1485
                    Height =300
                    Name ="Label69"
                    Caption ="Name"
                End
                Begin Line
                    BorderWidth =2
                    Top =2055
                    Width =9924
                    BorderColor =4210752
                    Name ="Line71"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =2880
                    Top =705
                    Width =4470
                    Height =660
                    FontSize =26
                    Name ="Label19"
                    Caption ="Chart Of Accounts"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextFontCharSet =161
                    TextAlign =3
                    TextFontFamily =18
                    Left =7995
                    Top =1046
                    Width =1964
                    Height =300
                    FontSize =9
                    FontWeight =700
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
                    Left =8218
                    Top =765
                    Width =1739
                    Height =256
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
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
                    TextAlign =2
                    TextFontFamily =18
                    Left =2895
                    Top =240
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
                Begin Line
                    BorderWidth =2
                    Top =1500
                    Width =9924
                    BorderColor =4210752
                    Name ="Line69"
                End
                Begin Label
                    Left =8475
                    Top =1695
                    Width =1155
                    Height =300
                    Name ="Label74"
                    Caption ="Account No."
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
                    CanGrow = NotDefault
                    TextFontCharSet =238
                    Left =1830
                    Top =45
                    Width =5520
                    Name ="coaName"
                    ControlSource ="coaName"

                End
                Begin TextBox
                    TextFontCharSet =238
                    Left =90
                    Top =45
                    Width =1665
                    TabIndex =1
                    Name ="coaRef"
                    ControlSource ="coaRef"

                End
                Begin Line
                    BorderWidth =1
                    Top =330
                    Width =9923
                    BorderColor =0
                    Name ="Line63"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7251
                    Top =47
                    TabIndex =2
                    Name ="lvlID"
                    ControlSource ="lvlID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8460
                    Top =45
                    TabIndex =3
                    Name ="accNo"
                    ControlSource ="accNo"

                End
            End
        End
        Begin PageFooter
            Height =307
            Name ="PageFooterSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =9924
                    BorderColor =0
                    Name ="Line36"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextAlign =2
                    TextFontFamily =18
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
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
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

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
  '  Me.coaRef = Me.coaRef
'    Me.coaNo.Move Left:=(Me.lvlID * 300)
    Me.coaName.Move Left:=(Me.lvlID * 100 + 1860)
    

End Sub

Private Sub Report_Open(Cancel As Integer)

If FnIsLoad("frmChart") Then
    Me.RecordSource = "qryChartOfAccounts5th"
ElseIf IsNull(Forms.frmReport.txtToStr1) = True Or Forms.frmReport.txtToStr1 = "" Then
    Me.RecordSource = "qryChartOfAccounts"
    Forms.frmReport.txtToStr1.Value = 5
End If

End Sub
