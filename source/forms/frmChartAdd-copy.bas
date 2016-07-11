Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
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
    Width =4677
    DatasheetFontHeight =10
    ItemSuffix =44
    Left =7215
    Top =2925
    Right =11895
    Bottom =4995
    DatasheetGridlinesColor =12632256
    Filter ="coaID=84"
    RecSrcDt = Begin
        0x76f45c25597ae340
    End
    RecordSource ="SELECT tblChart.* FROM tblChart; "
    Caption ="Chart Of Accounts - Add Category"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =1015
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    Left =240
                    Top =705
                    Width =1095
                    Height =225
                    ForeColor =128
                    Name ="crnShortName_Label"
                    Caption ="ShortName"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =135
                    Top =705
                    Width =660
                    Height =225
                    ForeColor =128
                    Name ="crnName_Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2985
                    Top =45
                    Width =240
                    Height =225
                    BorderColor =255
                    Name ="txtRefNo1"
                    ControlSource ="=Right(\"0\" & (CStr([coaRef1])),2)"
                    InputMask ="00"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1545
                            Top =45
                            Width =1395
                            Height =225
                            Name ="Label17"
                            Caption ="Current ref. no:"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3285
                    Top =45
                    Width =240
                    Height =225
                    TabIndex =1
                    BorderColor =255
                    Name ="txtRefNo2"
                    ControlSource ="=Right(\"0\" & (CStr([coaRef2])),2)"
                    InputMask ="00"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3585
                    Top =45
                    Width =240
                    Height =225
                    TabIndex =7
                    Name ="txtRefNo3"
                    ControlSource ="=Right(\"0\" & (CStr([coaRef3])),2)"
                    InputMask ="00"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3885
                    Top =45
                    Width =240
                    Height =225
                    TabIndex =8
                    Name ="txtRefNo4"
                    ControlSource ="=Right(\"0\" & (CStr([coaRef4])),2)"
                    InputMask ="00"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4170
                    Top =45
                    Width =390
                    Height =225
                    TabIndex =9
                    Name ="txtRefNo5"
                    ControlSource ="=Right(\"00\" & (CStr([coaRef5])),3)"
                    InputMask ="000"

                End
                Begin Label
                    OverlapFlags =247
                    Left =3210
                    Top =45
                    Width =120
                    Height =225
                    Name ="Label25"
                    Caption ="-"
                End
                Begin Label
                    OverlapFlags =247
                    Left =3510
                    Top =45
                    Width =120
                    Height =225
                    Name ="Label26"
                    Caption ="-"
                End
                Begin Label
                    OverlapFlags =247
                    Left =3810
                    Top =45
                    Width =120
                    Height =225
                    Name ="Label27"
                    Caption ="-"
                End
                Begin Label
                    OverlapFlags =247
                    Left =4095
                    Top =45
                    Width =120
                    Height =225
                    Name ="Label28"
                    Caption ="-"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3000
                    Top =570
                    Width =240
                    TabIndex =2
                    Name ="coaRef1"
                    ControlSource ="coaRef1"
                    DefaultValue ="0"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3285
                    Top =570
                    Width =240
                    TabIndex =3
                    Name ="coaRef2"
                    ControlSource ="coaRef2"
                    DefaultValue ="0"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3585
                    Top =570
                    Width =240
                    TabIndex =4
                    Name ="coaRef3"
                    ControlSource ="coaRef3"
                    DefaultValue ="0"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3855
                    Top =570
                    Width =240
                    TabIndex =5
                    Name ="coaRef4"
                    ControlSource ="coaRef4"
                    DefaultValue ="0"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4155
                    Top =570
                    Width =391
                    TabIndex =6
                    Name ="coaRef5"
                    ControlSource ="coaRef5"
                    DefaultValue ="0"

                End
                Begin Label
                    OverlapFlags =85
                    Left =1545
                    Top =570
                    Width =1395
                    Height =225
                    Name ="lblcoaRef"
                    Caption ="New ref. no:"
                End
                Begin Line
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3135
                    Top =300
                    Width =0
                    Height =213
                    Name ="lnRefNo1"
                End
                Begin Line
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3420
                    Top =300
                    Width =0
                    Height =213
                    Name ="lnRefNo2"
                End
                Begin Line
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3720
                    Top =300
                    Width =0
                    Height =213
                    Name ="lnRefNo3"
                End
                Begin Line
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =4005
                    Top =300
                    Width =0
                    Height =213
                    Name ="lnRefNo4"
                End
                Begin Line
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =4365
                    Top =300
                    Width =0
                    Height =213
                    Name ="lnRefNo5"
                End
            End
        End
        Begin Section
            Height =401
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =1095
                    Height =255
                    ColumnWidth =465
                    Name ="coaNo"
                    ControlSource ="coaNo"
                    Format =">"

                End
                Begin TextBox
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =90
                    Top =45
                    Width =3795
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="coaName"
                    ControlSource ="coaName"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Width =240
                    TabIndex =2
                    Name ="coaID"
                    ControlSource ="coaID"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3543
                    Top =23
                    Width =525
                    TabIndex =3
                    Name ="coaIDg"
                    ControlSource ="coaIDg"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =2763
                    Top =47
                    Width =495
                    Height =255
                    TabIndex =4
                    Name ="lvlID"
                    ControlSource ="lvlID"

                End
            End
        End
        Begin FormFooter
            Height =661
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =65
                    Left =1035
                    Top =195
                    Width =1080
                    Name ="cmdAdd"
                    Caption ="&Add New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add New Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =79
                    Left =2220
                    Top =195
                    Width =1080
                    TabIndex =1
                    Name ="cmdOK"
                    Caption ="&OK"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close Current Form"
                    UnicodeAccessKey =79

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =3450
                    Top =195
                    Width =1080
                    TabIndex =2
                    Name ="cmdCancel"
                    Caption ="&Cancel"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add New Record"
                    UnicodeAccessKey =67

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
Private Sub cmdAdd_Click()
'On Error GoTo Err_cmdAdd_Click
'  coaNo.Visible = True
'  coaName.Visible = True
'  Me.AllowAdditions = True
'  DoCmd.GoToRecord , , acNewRec
'  coaNo.SetFocus
'Exit_cmdAdd_Click:
'    Exit Sub
'Err_cmdAdd_Click:
'   MsgBox FnIsErr(Err.Number), vbExclamation
'    Resume Exit_cmdAdd_Click

End Sub

Private Sub cmdCancel_Click()
    If Left(Me.OpenArgs, 3) = "add" Then
        coaNo.Value = Empty
        coaRef1.Value = Empty
        coaRef2.Value = Empty
        coaRef3.Value = Empty
        coaRef4.Value = Empty
        coaRef5.Value = Empty
        lvlID.Value = Empty
        coaIDg.Value = Empty
    End If
    
    DoCmd.Close acForm, "frmChartAdd"
    Forms!frmChart.PopulateTreeView
 End Sub

Private Sub cmdOK_Click()
Dim strSource As String
    On Error GoTo Err_cmdOK_Click
    
    If Left(Me.OpenArgs, 3) = "add" Then
        Dim rs As Recordset
        Dim strSQL As String

        strSQL = "SELECT tblChart.coaID FROM tblChart WHERE tblChart.coaRef1 = " & coaRef1.Value & " AND tblChart.coaRef2 = " & coaRef2.Value & " AND tblChart.coaRef3 = " & coaRef3.Value & " AND tblChart.coaRef4 = " & coaRef4.Value & " AND tblChart.coaRef5 = " & coaRef5.Value

        Set rs = CurrentDb.OpenRecordset(strSQL)

        If rs.RecordCount > 0 Then
            MsgBox FnIsErr(3022), vbExclamation
            rs.Close
            Select Case lvlID
                Case 1
                    coaRef1.SetFocus
                Case 2
                    coaRef2.SetFocus
                Case 3
                    coaRef3.SetFocus
                Case 4
                coaRef4.SetFocus
                Case 5
                    coaRef5.SetFocus
            End Select
        Exit Sub
        End If
    End If
    DoCmd.Close acForm, "frmChartAdd"
    Forms!frmChart.PopulateTreeView
Exit_cmdOK_Click:
    Exit Sub
Err_cmdOK_Click:
   MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdOK_Click
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
    MsgBox FnIsErr(DataErr), vbExclamation, Err.Description
    Response = acDataErrContinue
End Sub

Private Sub Form_Load()
Dim tempcoaRef As String
On Error GoTo Err_cmdAdd_Click
  Me.Filter = ""
  Me.FilterOn = False
  If Left(Me.OpenArgs, 3) = "add" Then
    Me.AllowAdditions = True
    DoCmd.GoToRecord , , acNewRec
    lvlID.Value = Mid(Me.OpenArgs, 4, 1) + 1
    coaIDg.Value = Mid(Me.OpenArgs, 5, (InStr(6, Me.OpenArgs, "_", vbTextCompare) - 5))
    tempcoaRef = Mid(Me.OpenArgs, (InStr(6, Me.OpenArgs, "_", vbTextCompare) + 1), Len(Me.OpenArgs))
   
    Select Case lvlID
    Case 1
        coaRef1.Value = CDbl(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 9)) + 1
        coaIDg.Value = coaID.Value
        txtRefNo1.BorderWidth = 1
        txtRefNo1.BorderColor = 255
        txtRefNo1.BorderStyle = 1
        coaRef1.Visible = True
        lnRefNo1.Visible = True
    Case 2
        coaRef1.Value = CDbl(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 9))
        coaRef2.Value = CDbl(Right(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 7), 2)) + 1
        txtRefNo2.BorderWidth = 1
        txtRefNo2.BorderColor = 255
        txtRefNo2.BorderStyle = 1
        coaRef2.Visible = True
        lnRefNo2.Visible = True
    Case 3
        coaRef1.Value = CDbl(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 9))
        coaRef2.Value = CDbl(Right(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 7), 2))
        coaRef3.Value = CDbl(Right(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 5), 2)) + 1
        txtRefNo3.BorderWidth = 1
        txtRefNo3.BorderColor = 255
        txtRefNo3.BorderStyle = 1
        coaRef3.Visible = True
        lnRefNo3.Visible = True
    Case 4
        coaRef1.Value = CDbl(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 9))
        coaRef2.Value = CDbl(Right(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 7), 2))
        coaRef3.Value = CDbl(Right(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 5), 2))
        coaRef4.Value = CDbl(Right(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 3), 2)) + 1
        txtRefNo4.BorderWidth = 1
        txtRefNo4.BorderColor = 255
        txtRefNo4.BorderStyle = 1
        coaRef4.Visible = True
        lnRefNo4.Visible = True
    Case 5
        coaRef1.Value = CDbl(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 9))
        coaRef2.Value = CDbl(Right(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 7), 2))
        coaRef3.Value = CDbl(Right(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 5), 2))
        coaRef4.Value = CDbl(Right(Left(CStr(tempcoaRef), Len(CStr(tempcoaRef)) - 3), 2))
        coaRef5.Value = CDbl(Right(CStr(tempcoaRef), 3)) + 1
        txtRefNo5.BorderWidth = 1
        txtRefNo5.BorderColor = 255
        txtRefNo5.BorderStyle = 1
        coaRef5.Visible = True
        lnRefNo5.Visible = True
    End Select
    coaName.SetFocus
  Else
    Me.Filter = "coaID=" & Mid(Me.OpenArgs, 4, Len(Me.OpenArgs))
    Me.FilterOn = True
    lblcoaRef.Visible = False
  End If
Exit_cmdAdd_Click:
    Exit Sub
Err_cmdAdd_Click:
   MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdAdd_Click

End Sub
