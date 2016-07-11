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
    Width =7984
    DatasheetFontHeight =10
    ItemSuffix =16
    Left =7275
    Top =2715
    Right =16440
    Bottom =9030
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfd1ef60f2e99e340
    End
    RecordSource ="SELECT tblReference.refID, tblReference.refShortName, tblReference.refName, tblR"
        "eference.refInterFace FROM tblReference ORDER BY tblReference.refShortName; "
    Caption ="Reference File"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
        Begin FormHeader
            Height =1322
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
                    Caption ="Short Name"
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
                    TextFontCharSet =238
                    Left =255
                    Top =120
                    Width =4380
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="lblComent"
                    Caption ="Reference List"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =6375
                    Top =600
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
                    RowSource ="SELECT tblReference.refID, tblReference.refShortName, tblReference.refName FROM "
                        "tblReference; "
                    ColumnWidths ="0;567;2835"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =5955
                            Top =1035
                            Width =1905
                            Height =225
                            ForeColor =128
                            Name ="Label12"
                            Caption ="Assigned Document Type"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =255
                    Top =600
                    Width =585
                    Height =300
                    Name ="Label15"
                    Caption ="Search"
                End
            End
        End
        Begin Section
            Height =375
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =1095
                    Height =255
                    ColumnWidth =465
                    Name ="refShortName"
                    ControlSource ="refShortName"
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
                    Name ="refName"
                    ControlSource ="refName"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Width =240
                    TabIndex =2
                    Name ="refID"
                    ControlSource ="refID"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =5976
                    Top =60
                    Width =1871
                    Height =256
                    TabIndex =3
                    BorderColor =3
                    Name ="cboInterface"
                    ControlSource ="refInterFace"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Invoice\";2;\"Cash Sales\";3;\"Credit Note\";4;\"Debit Note\";5;\"Receipt\";"
                        "6;\"Expense\";7;\"Payment\";8;\"Deposit\";9;\"Withdraw\""
                    ColumnWidths ="0;1442"
                    OnExit ="[Event Procedure]"
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
                    AccessKey =77
                    Left =240
                    Top =240
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
                    ControlTipText ="First Record (Alt+m)"
                    UnicodeAccessKey =109
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =47
                    Left =1500
                    Top =240
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
                    Left =1080
                    Top =240
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
                    Left =660
                    Top =240
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
                    AccessKey =78
                    Left =4335
                    Top =240
                    Width =1080
                    TabIndex =4
                    Name ="cmdAdd"
                    Caption ="Add &New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add New Record (ALT+N)"
                    UnicodeAccessKey =78
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =5535
                    Top =240
                    Width =1080
                    TabIndex =5
                    Name ="cmdDel"
                    Caption ="&Delete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete Record (ALT+D)"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    Left =6735
                    Top =240
                    Width =1080
                    TabIndex =6
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close Current Form (ALT+X)"
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
    Me.RecordsetClone.FindFirst "[refID] = " & Me![cboFind]
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

Private Sub cboInterface_Exit(Cancel As Integer)
    Dim rst As Recordset
    Dim strSQL As String

'    strSQL = "SELECT tblAccount.accNo, tblAccount.accName FROM tblAccount WHERE tblAccount.caoID = " & caoID & " AND tblAccount.accNo <> '" & accNo & "'"
    If refID <> 0 And (IsNumeric(cboInterface) = True And cboInterface > 0) Then
        strSQL = "SELECT tblReference.refID, tblReference.refName FROM tblReference WHERE tblReference.refInterface = " & cboInterface & " AND tblReference.refID <> " & refID

        Set rst = CurrentDb.OpenRecordset(strSQL)
        If rst.RecordCount > 0 Then
            MsgBox ("This document is already assigned to the reference: " & rst.Fields(1))
            DoCmd.CancelEvent
        End If
    End If
    DoCmd.RunCommand acCmdSaveRecord
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
  refShortName.SetFocus
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
        refShortName.SetFocus
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
    DoCmd.Close
Else
    Select Case Me.OpenArgs
        Case "frmTransactionSub"
            DoCmd.Close acForm, "frmReference"
            Screen.ActiveForm.Refresh
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
