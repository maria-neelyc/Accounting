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
    ItemSuffix =37
    Left =1005
    Top =255
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    Filter ="CR Is Null And DR Is Null AND CInt(lvlID) = 5"
    OrderBy ="coaRef1, coaRef2, coaRef3, coaRef4, coaRef5"
    Toolbar ="PrintReportEmail"
    RecSrcDt = Begin
        0x1154dcb27985e340
    End
    RecordSource ="qryBalanceSheet3"
    Caption ="Balance Sheet"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x27010000370200003702000037020000000000004b2a00000801000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =238
            TextAlign =1
            FontSize =10
            FontWeight =700
            FontName ="Arial CE"
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
            BorderWidth =2
            Width =850
            Height =850
            BorderColor =12632256
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
            Width =1701
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
            Width =1701
            Height =1701
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
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
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
            TextFontCharSet =238
            TextFontFamily =18
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman CE"
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
            TextFontCharSet =238
            TextFontFamily =18
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman CE"
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
            TextFontCharSet =238
            TextFontFamily =18
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman CE"
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
            Width =1701
            Height =1701
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
        Begin UnboundObjectFrame
            OldBorderStyle =1
            Width =4536
            Height =2835
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
        Begin BreakLevel
            ControlSource ="coaFRef"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =2175
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1485
                    Top =1650
                    Width =3000
                    Height =270
                    Name ="coaName_Label"
                    Caption ="Ref. Heading"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
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
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Top =1650
                    Width =1425
                    Height =270
                    Name ="coaRef_Label"
                    Caption ="Reference"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
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
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =4530
                    Top =1650
                    Width =4432
                    Height =270
                    Name ="accName_Label"
                    Caption ="Account"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
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
                    BorderWidth =2
                    Top =2160
                    Width =10827
                    BorderColor =4210752
                    Name ="Line25"
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
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =8970
                    Top =1650
                    Width =885
                    Height =270
                    Name ="Label28"
                    Caption ="Debits"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
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
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =9885
                    Top =1650
                    Width =885
                    Height =270
                    Name ="Label29"
                    Caption ="Credits"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
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
                    Caption ="Balance Sheet"
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
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =6015
                    Top =1260
                    Width =1200
                    Height =270
                    Name ="lblPrdFrom"
                    Caption ="Period From:"
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
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =8385
                    Top =1260
                    Width =1260
                    Height =270
                    Name ="lblPrdTo"
                    Caption ="To:"
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
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7260
                    Top =1260
                    Width =1086
                    FontSize =8
                    TabIndex =4
                    Name ="PeriodFrom"
                    ControlSource ="=Forms!frmReport!txtFromStr1 & ' (' & IIf(Forms!frmReport!txtFromStr1=0,0,MonthN"
                        "ame(Forms!frmReport!txtFromStr1)) & ')'"
                    FontName ="Arial"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =9705
                    Top =1260
                    Width =1086
                    FontSize =8
                    TabIndex =5
                    Name ="PeriodTo"
                    ControlSource ="=Forms!frmReport!txtToStr1 & ' (' & IIf(Forms!frmReport!txtToStr1=0,0,MonthName("
                        "Forms!frmReport!txtToStr1)) & ')'"
                    FontName ="Arial"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =264
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1485
                    Width =2880
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
                    Left =4470
                    Width =1140
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
                    Left =5670
                    Width =3247
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
                    Left =9900
                    Width =851
                    Height =260
                    FontSize =8
                    TabIndex =4
                    Name ="txtCR"
                    ControlSource ="=Abs(IIf((IIf(IsNumeric([DR]),[DR]+[OpenBalDR],[OpenBalDR])-IIf(IsNumeric([CR]),"
                        "[CR]-[OpenBalCR],-[OpenBalCR]))<0,(IIf(IsNumeric([DR]),[DR]+[OpenBalDR],[OpenBal"
                        "DR])-IIf(IsNumeric([CR]),[CR]-[OpenBalCR],-[OpenBalCR]))))"
                    Format ="Standard"
                    FontName ="Arial"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8985
                    Width =851
                    Height =260
                    FontSize =8
                    TabIndex =5
                    Name ="txtDR"
                    ControlSource ="=IIf((IIf(IsNumeric([DR]),[DR]+[OpenBalDR],[OpenBalDR])-IIf(IsNumeric([CR]),[CR]"
                        "-[OpenBalCR],-[OpenBalCR]))>0,(IIf(IsNumeric([DR]),[DR]+[OpenBalDR],[OpenBalDR])"
                        "-IIf(IsNumeric([CR]),[CR]-[OpenBalCR],-[OpenBalCR])))"
                    Format ="Standard"
                    FontName ="Arial"

                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8390
                    Width =451
                    Height =260
                    FontSize =8
                    TabIndex =6
                    Name ="lvlID"
                    ControlSource ="lvlID"
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
                    Left =8970
                    Top =60
                    Width =851
                    Height =260
                    FontSize =8
                    Name ="Text31"
                    ControlSource ="=Sum(IIf((IIf(IsNumeric([DR]),[DR]+[OpenBalDR],[OpenBalDR])-IIf(IsNumeric([CR]),"
                        "[CR]-[OpenBalCR],-[OpenBalCR]))>0,(IIf(IsNumeric([DR]),[DR]+[OpenBalDR],[OpenBal"
                        "DR])-IIf(IsNumeric([CR]),[CR]-[OpenBalCR],-[OpenBalCR]))))"
                    Format ="Standard"
                    FontName ="Arial"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9885
                    Top =60
                    Width =851
                    Height =260
                    FontSize =8
                    TabIndex =1
                    Name ="Text33"
                    ControlSource ="=Sum(Abs(IIf((IIf(IsNumeric([DR]),[DR]+[OpenBalDR],[OpenBalDR])-IIf(IsNumeric([C"
                        "R]),[CR]-[OpenBalCR],-[OpenBalCR]))<0,(IIf(IsNumeric([DR]),[DR]+[OpenBalDR],[Ope"
                        "nBalDR])-IIf(IsNumeric([CR]),[CR]-[OpenBalCR],-[OpenBalCR])))))"
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

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
Dim intBold, intNorm As Integer

intBold = 700
intNorm = 400
    
    If lvlID = 1 Then
        coaRef.FontWeight = intBold
        coaName.FontWeight = intBold
        accNo.FontWeight = intBold
        accName.FontWeight = intBold
        txtDR.FontWeight = intBold
        txtCR.FontWeight = intBold
    Else
        coaRef.FontWeight = intNorm
        coaName.FontWeight = intNorm
        accNo.FontWeight = intNorm
        accName.FontWeight = intNorm
        txtDR.FontWeight = intNorm
        txtCR.FontWeight = intNorm
    End If
End Sub

Private Sub Report_NoData(Cancel As Integer)

MsgBox ("The report is empty.")
Cancel = True

End Sub
Private Sub Report_Open(Cancel As Integer)
If IsNull(Forms.frmReport.txtFromStr1) = True Or Forms.frmReport.txtFromStr1 = "" Then
    Forms.frmReport.txtFromStr1.Value = 0
End If
If IsNull(Forms.frmReport.txtToStr1) = True Or Forms.frmReport.txtToStr1 = "" Then
    Forms.frmReport.txtToStr1.Value = 12
End If
If IsNull(Forms.frmReport.txtToStr2) = True Or Forms.frmReport.txtToStr2 = "" Then
    intCurLev = 5
    strFromRef = Forms.frmReport.FromRef.Value
    Do Until (intCurLev = 0)
        If (intCurLev = 5) Then
            intSub = CInt(Right(strFromRef, 3))
            strFromRef = Left(strFromRef, Len(strFromRef) - 3)
        Else
            intSub = CInt(Right(strFromRef, 2))
            strFromRef = Left(strFromRef, Len(strFromRef) - 2)
        End If
        If (intSub > 0) Then
            Forms.frmReport.txtToStr2.Value = intCurLev
            Exit Do
        End If
        intCurLev = intCurLev - 1
    Loop
    If (intCurLev = 0) Then
        Forms.frmReport.txtToStr2.Value = 5
    End If

End If

Forms.frmReport.ToRef.Value = Left(Forms.frmReport.ToRef.Value & "99999999999", 11)

If Forms.frmReport.cboOrderBy = 1 Then
    Me.OrderBy = "coaRef"
Else
    Me.OrderBy = "accNo"
End If

If Forms.frmReport.txtFromStr1.Value = 0 Then
    lblPrdFrom.Visible = False
    PeriodFrom.Visible = False
    lblPrdTo.Caption = "Period Up To:"
Else
    lblPrdFrom.Visible = True
    PeriodFrom.Visible = True
    lblPrdTo.Caption = "To:"
End If


Select Case Forms.frmReport.grpOptions
    Case 1
        Me.Filter = "((Bal = 0) OR (IsNumeric(Bal)=False))  AND CInt(lvlID) = 5 "
        Me.FilterOn = True
    Case 2
        Me.Filter = "CR Is Null And DR Is Null AND CInt(lvlID) = 5"
        Me.FilterOn = True
    Case 3
        Me.Filter = "((CR >0 Or DR > 0) AND CInt(lvlID) = 5)" ' OR CInt(lvlID) < 5
        Me.FilterOn = True
    Case 4
        Me.Filter = " (Bal <> 0 AND CInt(lvlID) = 5) " 'OR CInt(lvlID) < 5
        Me.FilterOn = True
    Case 5
        Me.Filter = ""
        Me.FilterOn = False
    Case Else
        Me.Filter = ""
        Me.FilterOn = False
End Select
    


End Sub
