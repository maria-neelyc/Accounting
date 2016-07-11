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
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6180
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =3450
    Top =2550
    Right =10650
    Bottom =8865
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe15fc9bbf46ee340
    End
    RecordSource ="SELECT tblDescription.desID, tblDescription.desShortName, tblDescription.desName"
        " FROM tblDescription ORDER BY tblDescription.desShortName; "
    Caption ="Reference File"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
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
            SpecialEffect =3
            BackStyle =0
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
            BackStyle =0
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
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
        Begin OptionButton
            SpecialEffect =2
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
        Begin CheckBox
            SpecialEffect =2
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
        Begin OptionGroup
            SpecialEffect =3
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
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
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
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
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
            ShowDatePicker =1
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
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
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
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
            SpecialEffect =2
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
            SpecialEffect =2
            OldBorderStyle =1
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
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
        Begin Tab
            BackStyle =0
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
            Height =1275
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =240
                    Top =1035
                    Width =1095
                    Height =225
                    ForeColor =128
                    Name ="crnShortName_Label"
                    Caption ="ShortName"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1500
                    Top =1035
                    Width =4260
                    Height =225
                    ForeColor =128
                    Name ="crnName_Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    Left =921
                    Top =118
                    Width =4380
                    Height =420
                    Name ="lblComent"
                    Caption ="List of References which are used in Entry Form.\015\012To fast search use provi"
                        "ded Combo-Box."
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =3861
                    Top =598
                    Width =1440
                    Height =240
                    FontWeight =700
                    Name ="lblPos"
                    Caption ="Rec. New"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3402
                    Left =966
                    Top =598
                    Width =2880
                    Height =300
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"6\""
                    Name ="cboFind"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblDescription.desID, tblDescription.desShortName, tblDescription.desName"
                        " FROM tblDescription ORDER BY tblDescription.desShortName; "
                    ColumnWidths ="0;567;2835"
                    AfterUpdate ="[Event Procedure]"
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =1095
                    Height =255
                    ColumnWidth =465
                    Name ="desShortName"
                    ControlSource ="desShortName"
                    Format =">"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =60
                    Width =4260
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="desName"
                    ControlSource ="desName"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Width =240
                    TabIndex =2
                    Name ="desID"
                    ControlSource ="desID"

                End
            End
        End
        Begin FormFooter
            Height =660
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =70
                    Left =240
                    Top =240
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

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =76
                    Left =1500
                    Top =240
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

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    Left =1080
                    Top =240
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

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =80
                    Left =660
                    Top =240
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

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    Left =2640
                    Top =240
                    Width =1080
                    TabIndex =4
                    Name ="cmdAdd"
                    Caption ="&Add New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add New Record"

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =3840
                    Top =240
                    Width =1080
                    TabIndex =5
                    Name ="cmdDel"
                    Caption ="&Delete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete Record"

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    Left =5040
                    Top =240
                    Width =1080
                    TabIndex =6
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close Current Form"

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
Private Sub cboFind_AfterUpdate()
On Error GoTo Err_End
' Find the record that matches the control.
If IsNoData(cboFind.Value) = False Then
    Me.Requery
    Me.RecordsetClone.FindFirst "[desID] = " & Me![cboFind]
    Me.Bookmark = Me.RecordsetClone.Bookmark
End If
Exit Sub
Err_End:
        Dim strMsg As String
        strMsg = FnIsErr(Err.Number)
        If strMsg <> "" Or IsNull(strMsg) Or IsEmpty(strMsg) Then
            MsgBox strMsg, vbExclamation, "Error!"
        Else
        End If
        cboFind.Value = ""
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
  desShortName.SetFocus
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
        desShortName.SetFocus
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
        Case "frmTransactionSingle"
            DoCmd.Close acForm, Me.Form.name
            Screen.ActiveForm.Refresh
            Forms.frmTransactionSingle.Requery
        Case Else
    End Select
End If
End Sub
Private Sub Form_Current()
' Purpose: Show current record position
If Me.NewRecord Then
    Me!lblPos.Caption = "Rec. New"
Else
    Me.RecordsetClone.Bookmark = Me.Bookmark
    Me!lblPos.Caption = "Rec. " & CStr(Me.RecordsetClone.AbsolutePosition + 1) _
                        & " of " & CStr(Me.RecordsetClone.RecordCount)
End If
cboFind.Requery
End Sub
