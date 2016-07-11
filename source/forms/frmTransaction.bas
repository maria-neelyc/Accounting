Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =14787
    DatasheetFontHeight =10
    ItemSuffix =112
    Left =2325
    Top =360
    Right =17115
    Bottom =7950
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5b2122297789e340
    End
    RecordSource ="SELECT tblTransaction.* FROM tblTransaction ORDER BY tblTransaction.trnID; "
    Caption ="=FnIME"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnError ="[Event Procedure]"
    OnDataChange ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            BackStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =1086
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =11475
                    Top =90
                    Width =1920
                    Height =300
                    FontWeight =700
                    Name ="lblPos"
                    Caption ="Rec. 1 of 1"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =3402
                    Left =1686
                    Top =330
                    Width =2235
                    Height =285
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cboFind"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblTransaction.trnID, tblTransaction.trnPageCounter, tblTransaction.trnIn"
                        "ternalRef, tblTransaction.trnEntryDate FROM tblTransaction ORDER BY tblTransacti"
                        "on.trnPageCounter; "
                    ColumnWidths ="0;1134;1134;1134"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial Greek"
                    OnGotFocus ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =810
                            Top =334
                            Width =795
                            Height =210
                            Name ="Label91"
                            Caption ="Find Page"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =255
                    Left =165
                    Top =45
                    Width =3969
                    Height =688
                    TabIndex =1
                    Name ="grpFind"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =708
                            Top =45
                            Width =960
                            Height =225
                            FontWeight =700
                            ForeColor =0
                            Name ="Label150"
                            Caption ="Search By"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    TextFontCharSet =161
                    ColumnCount =2
                    ListWidth =2880
                    Left =1677
                    Top =45
                    Width =1418
                    Height =260
                    TabIndex =2
                    Name ="cboSearch"
                    RowSourceType ="Value List"
                    RowSource ="\"Page No\";1;\"Voucher\";2;\"Date\";3"
                    ColumnWidths ="1442;0"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    DefaultValue ="\"Page No\""
                    FontName ="Arial Greek"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =9045
                    Top =90
                    Width =2040
                    Height =270
                    ForeColor =255
                    Name ="lbllock"
                    Caption ="Page Is Locked. Read Only"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6432
                    Top =91
                    TabIndex =3
                    BackColor =-2147483633
                    ForeColor =16711680
                    Name ="trnInterFace"
                    ControlSource ="trnInterface"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4380
                            Top =90
                            Width =1995
                            Height =240
                            ForeColor =16711680
                            Name ="lblInterFace"
                            Caption ="Transaction was issued by:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =2730
                    Top =810
                    Width =1050
                    Height =255
                    Name ="Label111"
                    Caption ="Last Voucher"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5640
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1162
                    Top =354
                    Width =2235
                    Height =255
                    TabIndex =5
                    Name ="trnEntryDate"
                    ControlSource ="trnEntryDate"
                    Format ="dd\\/mm\\/yyyy"
                    DefaultValue ="=Date()"
                    InputMask ="00/00/0000;0;_"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =637
                            Top =354
                            Width =465
                            Height =255
                            Name ="bnkaABA_Label"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1155
                    Top =75
                    Width =1380
                    Height =255
                    TabIndex =1
                    Name ="trnInternalRef"
                    ControlSource ="trnInternalRef"
                    FontName ="Arial Greek"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =240
                            Top =75
                            Width =855
                            Height =255
                            Name ="bnkaBLZ_Label"
                            Caption ="Voucher"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1191
                    Left =4980
                    Top =58
                    Width =960
                    Height =270
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="persID"
                    ControlSource ="persID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblPeriodSel.persID, tblPeriod.perName, tblPeriodSel.perID FROM tblPeriod"
                        "Sel INNER JOIN tblPeriod ON tblPeriodSel.perID=tblPeriod.perID ORDER BY tblPerio"
                        "dSel.persID; "
                    ColumnWidths ="0;1134;0"
                    OnExit ="[Event Procedure]"
                    DefaultValue ="Month(Date())"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4140
                            Top =58
                            Width =780
                            Height =255
                            Name ="crnID_Label"
                            Caption ="Period"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4995
                    Top =373
                    Width =2535
                    Height =255
                    TabIndex =6
                    Name ="trnPageCounter"
                    ControlSource ="trnPageCounter"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4410
                            Top =373
                            Width =525
                            Height =255
                            Name ="bnkaIBAN_Label"
                            Caption ="Page"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    Width =255
                    Name ="trnID"
                    ControlSource ="trnID"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5985
                    Top =58
                    Width =450
                    Height =255
                    TabIndex =4
                    Name ="Text80"
                    ControlSource ="persID"
                    DefaultValue ="Month(Date())"
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =90
                    Top =660
                    Width =14655
                    Height =4980
                    TabIndex =7
                    Name ="frmSub"
                    SourceObject ="Form.frmTransactionSub"
                    LinkChildFields ="trnID"
                    LinkMasterFields ="trnID"
                    OnEnter ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListWidth =1191
                    Left =6480
                    Top =45
                    Width =1050
                    Height =270
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"4\";\"4\""
                    Name ="trnYear"
                    ControlSource ="trnYear"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblControl.yearID FROM tblControl; "
                    OnExit ="[Event Procedure]"
                    DefaultValue ="=Year(Date())"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =8421
                    Top =328
                    Width =285
                    TabIndex =8
                    Name ="trnLock"
                    ControlSource ="trnLock"
                    StatusBarText ="0 - no lock; 1- lock;"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =9342
                    Top =352
                    TabIndex =9
                    Name ="PrdChk"
                    DefaultValue ="0"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =10405
                    Top =287
                    TabIndex =10
                    Name ="HasChanged"
                    DefaultValue ="0"
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2574
                    Top =75
                    TabIndex =11
                    Name ="lastVoucher"
                    ControlSource ="=DLookUp(\"trnInternalRef\",\"tblTransaction\",\"trnID = \" & NZ(DLookUp(\"Max(t"
                        "rnID)\",\"tblTransaction\"),0))"
                End
            End
        End
        Begin FormFooter
            Height =885
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =77
                    Left =120
                    Top =60
                    Width =366
                    Name ="cmdFirst"
                    Caption ="&m"
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="First Record (Alt+M)"
                    UnicodeAccessKey =109
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =47
                    Left =1380
                    Top =60
                    Width =366
                    TabIndex =1
                    Name ="cmdLast"
                    Caption ="&/"
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Last Record (Alt+/)"
                    UnicodeAccessKey =47
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =46
                    Left =960
                    Top =60
                    Width =366
                    TabIndex =2
                    Name ="cmdNext"
                    Caption ="&."
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Next Record (Alt+.)"
                    UnicodeAccessKey =46
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =44
                    Left =540
                    Top =60
                    Width =366
                    TabIndex =3
                    Name ="cmdPrevious"
                    Caption ="&,"
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Previous Record (Alt+,)"
                    UnicodeAccessKey =44
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =87
                    Left =6450
                    Top =45
                    Width =1080
                    TabIndex =4
                    Name ="cmdAdd"
                    Caption ="Ne&w Page"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add New Record (ALT+N)"
                    UnicodeAccessKey =119
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    Left =12268
                    Top =45
                    Width =1080
                    TabIndex =5
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close Current Form (ALT+X)"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =83
                    Left =9986
                    Top =45
                    Width =1080
                    TabIndex =6
                    Name ="cmdSave"
                    Caption ="&Save"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save Page (ALT+S)"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    Left =8831
                    Top =45
                    Width =1080
                    TabIndex =7
                    Name ="cmdEdit"
                    Caption ="&Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Edit Page (ALT+E)"
                    UnicodeAccessKey =69
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =81
                    Left =7650
                    Top =45
                    Width =1080
                    TabIndex =8
                    Name ="cmdUndo"
                    Caption ="&Quick exit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Exits transaction details if none exists  (ALT+Q)"
                    UnicodeAccessKey =81
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =82
                    Left =11130
                    Top =45
                    Width =1080
                    TabIndex =9
                    Name ="cmdPrint"
                    Caption ="P&rint"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Print transactions of this page (ALT+R)"
                    UnicodeAccessKey =114
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
'On Error GoTo Err_End
' Find the record that matches the control.
If IsNoData(cboFind.Value) = False Then
    Me.Requery
    Me.RecordsetClone.FindFirst "[trnID] = " & Me![cboFind]
    Me.Bookmark = Me.RecordsetClone.Bookmark
    Me.AllowEdits = False
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
        Me.AllowEdits = False
End Sub

Private Sub cboFind_GotFocus()
Me.AllowEdits = True
End Sub

Private Sub cboSearch_AfterUpdate()
'cboFind.InputMask = "99/99/00;0;_"

Select Case cboSearch.Column(1)
    Case 1
    cboFind.InputMask = ""
    cboFind.RowSource = " SELECT tblTransaction.trnID, tblTransaction.trnPageCounter, tblTransaction.trnInternalRef, tblTransaction.trnEntryDate " & _
              " FROM tblTransaction ORDER BY tblTransaction.trnPageCounter;"


    Case 2
        cboFind.InputMask = ""
        cboFind.RowSource = " SELECT tblTransaction.trnID, tblTransaction.trnInternalRef, tblTransaction.trnPageCounter, tblTransaction.trnEntryDate " & _
              " FROM tblTransaction ORDER BY tblTransaction.trnInternalRef, tblTransaction.trnPageCounter;"

    Case 3
        cboFind.InputMask = "99/99/00;0;_"
        cboFind.RowSource = " SELECT tblTransaction.trnID, tblTransaction.trnEntryDate, tblTransaction.trnPageCounter, tblTransaction.trnInternalRef  " & _
              " FROM tblTransaction ORDER BY tblTransaction.trnEntryDate, tblTransaction.trnPageCounter;"
End Select
End Sub

Private Sub cboSearch_Enter()
    Me.AllowEdits = True
End Sub

Private Sub cboSearch_Exit(Cancel As Integer)
    grpFind.DefaultValue = cboSearch.Column(1)
    
    Me.AllowEdits = False
End Sub


Private Sub cmdEdit_Click()
On Error GoTo Err_cmdEdit_Click

If trnLock = 1 Then
    MsgBox "You can't edit this page. Page is locked.", vbInformation
    Exit Sub
End If

  Me.AllowEdits = True
  frmSub.Enabled = True
  frmSub.Form.AllowEdits = True
  frmSub.Form.AllowAdditions = True
  frmSub.Form.cmdTrns.Enabled = True
  trnInternalRef.SetFocus
  cmdSave.Enabled = True
  cmdExit.Enabled = False
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdUndo.Enabled = True
    cmdFirst.Enabled = False
  cmdPrevious.Enabled = False
  cmdNext.Enabled = False
  cmdLast.Enabled = False
   cboFind.Enabled = False
  
Exit_cmdEdit_Click:
    Exit Sub
Err_cmdEdit_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdEdit_Click
End Sub

Private Sub cmdFirst_Click()
On Error GoTo Err_cmdFirst_Click
'If frmSub.Enabled = True Then
'    frmSub.Controls("accID1").SetFocus
'    If Save = False Then
'        Exit Sub
'    End If
'End If

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
'If frmSub.Enabled = True Then
'    frmSub.Controls("accID1").SetFocus
'    If Save = False Then
'        Exit Sub
'    End If
'End If
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
'If frmSub.Enabled = True Then
'    frmSub.Controls("accID1").SetFocus
'    If Save = False Then
'        Exit Sub
'    End If
'End If

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
'If frmSub.Enabled = True Then
'    frmSub.Controls("accID1").SetFocus
'    If Save = False Then
'        Exit Sub
'    End If
'End If

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
  Me.AllowEdits = True
  
'If frmSub.Enabled = True Then
'    frmSub.Controls("accID1").SetFocus
'End If
  
  DoCmd.GoToRecord , , acNewRec
  frmSub.Enabled = True
  frmSub.Form.AllowEdits = True
  frmSub.Form.AllowAdditions = True
  frmSub.Form.cmdTrns.Enabled = True
  trnInternalRef.SetFocus
  cmdSave.Enabled = True
  cmdExit.Enabled = False
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdUndo.Enabled = True
    cmdFirst.Enabled = False
  cmdPrevious.Enabled = False
  cmdNext.Enabled = False
  cmdLast.Enabled = False
   cboFind.Enabled = False
  
  
  
'  trnInternalRef.SetFocus
'  cmdSave.Enabled = True
'  cmdExit.Enabled = False
'  cmdAdd.Enabled = False
'  cmdEdit.Enabled = False
'  cmdUndo.Enabled = False
'    cmdFirst.Enabled = False
'  cmdPrevious.Enabled = False
'  cmdNext.Enabled = False
'  cmdLast.Enabled = False
'  frmSub.Form.cmdTrns.Enabled = True
'   cboFind.Enabled = False
   
  If IsNoData(trnPageCounter) = True Then
        Dim InternalCounter As New clsCounter
        InternalCounter.Location "tblControl", "ctrPageCounter"
        InternalCounter.Value = InternalCounter.Value + 1
        trnPageCounter.Value = InternalCounter.Value
        'trnPageCounter.Enabled = True
        'trnPageCounter.Locked = True
  End If
  
Exit_cmdAdd_Click:
    Exit Sub
Err_cmdAdd_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdAdd_Click
End Sub

Private Sub cmdExit_Click()
If Me.Recordset.RecordCount > 0 Then
    If frmSub.Enabled = True And Me.frmSub.Form.Recordset.RecordCount > 0 Then
        frmSub.Controls("accID1").SetFocus
    End If
End If

'Automatic update of next year's opening balances if any changes has been made for
'current year's treansactions and if next year exists
'    If fsClosedYear And HasChanged Then
'        If fsNextYearHasChanged Then
'            MsgBox FnIsErr(7009), vbExclamation
'        End If
'        sUpdtBalances
'        HasChanged = False
'    End If
    
    DoCmd.OpenForm "frmMain"
    DoCmd.Close acForm, "frmTransaction"
End Sub

Private Sub cmdPrint_Click()
DoCmd.OpenReport "rptTrnsPages", acViewPreview
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_cmdSave_Click
If Save = False Then
    Exit Sub
End If
  cmdFirst.Enabled = True
  cmdPrevious.Enabled = True
  cmdNext.Enabled = True
  cmdLast.Enabled = True
   cboFind.Enabled = True
  
Exit_cmdSave_Click:
    cmdUndo.Enabled = False
    Exit Sub
Err_cmdSave_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdSave_Click
End Sub

Private Sub cmdUndo_Click()
On Error GoTo Err_cmdUndo_Click

    If DLookup("Count(trnsID)", "tblTransactionSub", "trnID = " & trnID) = 0 Then
        'Dim intTrnID As Integer
        'intTrnID = trnID
        'DoCmd.SetWarnings False
        'DoCmd.RunSQL ("DELETE FROM tblTransaction WHERE trnID = " & intTrnID)
        'DoCmd.SetWarnings True
        'DoCmd.GoToRecord , , acPrevious
        
        frmSub.Enabled = False
        frmSub.Form.AllowEdits = False
        frmSub.Form.AllowAdditions = False
        frmSub.Form.cmdTrns.Enabled = False
        cmdSave.Enabled = False
        cmdExit.Enabled = True
        cmdAdd.Enabled = True
        cmdEdit.Enabled = True
        cmdEdit.SetFocus
        cmdUndo.Enabled = False
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        cboFind.Enabled = True
    Else
        Exit Sub
    End If

'  DoCmd.RunCommand acCmdUndo
'  Me.cmdUndo.Enabled = False
Exit_cmdUndo_Click:
    Exit Sub
Err_cmdUndo_Click:
    If Err.Number = 2046 Then
        Resume Exit_cmdUndo_Click
    End If
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdUndo_Click
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
Me.Caption = FnIme()
cmdEdit.Enabled = Not Me.trnLock
lbllock.Visible = trnLock
If trnInterface <> "" Then
    trnInterface.Visible = True
    lblInterFace.Visible = True
Else
    trnInterface.Visible = False
    lblInterFace.Visible = False
End If

cboFind.Requery
End Sub

Private Sub Form_DataChange(ByVal Reason As Long)
HasChanged = True
End Sub



Private Sub Form_Error(DataErr As Integer, Response As Integer)
Dim strMsg As String
        strMsg = FnIsErr(DataErr)
        Response = acDataErrContinue
        If IsNoData(strMsg) = False Then
            MsgBox strMsg, vbExclamation, "Error!"
        Else
        End If
End Sub

Private Sub Form_Open(Cancel As Integer)
HasChanged = False
End Sub
Private Sub frmSub_Enter()
    'DoCmd.RunCommand acCmdSave
'    Me.Refresh
End Sub

Private Sub persID_Exit(Cancel As Integer)
If (persID <> DLookup("persID", "Current_Period")) And (PrdChk <> True) Then
    MsgBox "The transaction period is different than current period", vbOKOnly, "Warning"
    PrdChk = True
End If
End Sub

Private Sub trnsSort_Click()
    If trnsSort.Caption = "&Code" Then
        Me.frmSub.Form.accID.RowSource = "SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName, tblAccount.accSign " _
                        & "FROM tblAccount " _
                        & "WHERE (((tblAccount.accStatus) = True)) " _
                        & "ORDER BY tblAccount.accNo;"
        Me.frmSub.Form.accID.ColumnWidths = "0cm;1,507cm;4cm;0cm"
        'Me.trnsSort.Caption = "Na&me"
    Else
        Me.frmSub.Form.accID.RowSource = "SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo, tblAccount.accSign " _
                        & "FROM tblAccount " _
                        & "WHERE (((tblAccount.accStatus) = True)) " _
                        & "ORDER BY tblAccount.accName;"
        Me.frmSub.Form.accID.ColumnWidths = "0cm;4cm;1,507cm;0cm"
        'Me.trnsSort.Caption = "Nu&mber"
    End If
    Me.frmSub.Form.accID.Requery
End Sub


Private Sub trnYear_Exit(Cancel As Integer)
If (persID <> DLookup("persID", "Current_Period")) And (PrdChk <> True) Then
    MsgBox "The transaction period is different than current period", vbOKOnly, "Warning"
    PrdChk = True
End If
End Sub

Private Function Save() As Boolean
frmSub.Controls("txtSumAmount").Requery
Dim dblSumCR, dblSumDR As Double

dblSumDR = RoundDG(IIf(IsNumeric(fsCR_DB([trnID], "trnsDebits")), fsCR_DB([trnID], "trnsDebits"), 0))
dblSumCR = RoundDG(IIf(IsNumeric(fsCR_DB([trnID], "trnsCredits")), fsCR_DB([trnID], "trnsCredits"), 0))

'If frmSub.Controls("txtSumAmount") <> 0 Then
If (dblSumDR - dblSumCR) <> 0 Then
    MsgBox "Current Debit and Credit transactions do not match", vbOKOnly, "Incorrect Balance"
    Save = False
    Exit Function
End If
  
If IsNull(Me.trnInternalRef) Then
    MsgBox ("Cannot save. Please fill in the vaucher no.")
    Me.trnInternalRef.SetFocus
    Save = False
    Exit Function
End If
  
If frmSub.Enabled = True Then
    frmSub.Controls("accID1").SetFocus
End If
DoCmd.RunCommand acCmdSaveRecord
Me.AllowAdditions = False
frmSub.Form.AllowAdditions = False
Me.AllowEdits = False
frmSub.Form.AllowEdits = False
frmSub.Form.cmdTrns.Enabled = False
frmSub.Enabled = False
trnInternalRef.SetFocus
cmdSave.Enabled = False
cmdExit.Enabled = True
cmdAdd.Enabled = True
cmdEdit.Enabled = True
cboFind.Requery
Save = True

End Function
