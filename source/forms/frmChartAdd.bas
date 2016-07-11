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
    FilterOn = NotDefault
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
    Width =4700
    DatasheetFontHeight =10
    ItemSuffix =49
    Left =7200
    Top =2790
    Right =11895
    Bottom =5190
    DatasheetGridlinesColor =12632256
    Filter ="coaID=129"
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
        Begin Line
            Width =1701
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
            Height =1015
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
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
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =377
                    Top =283
                    TabIndex =10
                    Name ="HasChanged"
                    DefaultValue ="0"
                End
            End
        End
        Begin Section
            Height =732
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1620
                    Top =390
                    Width =1095
                    Height =255
                    ColumnWidth =465
                    Name ="coaNo"
                    ControlSource ="coaNo"
                    Format =">"
                End
                Begin TextBox
                    OverlapFlags =255
                    TextFontCharSet =161
                    IMESentenceMode =3
                    Left =1380
                    Top =375
                    Width =3105
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="coaName"
                    ControlSource ="coaName"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial Greek"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1380
                    Top =330
                    Width =240
                    TabIndex =3
                    Name ="coaID"
                    ControlSource ="coaID"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3466
                    Top =401
                    Width =525
                    TabIndex =4
                    Name ="coaIDg"
                    ControlSource ="coaIDg"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =4053
                    Top =377
                    Width =495
                    Height =255
                    TabIndex =5
                    Name ="lvlID"
                    ControlSource ="lvlID"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =161
                    IMESentenceMode =3
                    Left =1380
                    Top =45
                    Width =3105
                    Height =255
                    TabIndex =1
                    Name ="coaNameEng"
                    ControlSource ="coaNameEng"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial Greek"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =30
                    Top =75
                    Width =660
                    Height =225
                    ForeColor =128
                    Name ="Label47"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =30
                    Top =375
                    Width =1290
                    Height =225
                    ForeColor =128
                    Name ="lblTransl"
                    Caption ="Translated Name"
                    Tag ="DetachedLabel"
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
                    ControlTipText ="Add New Record In the same parent level  (ALT+A)"
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
                    ControlTipText ="Save and exit  (ALT+O)"
                    UnicodeAccessKey =79
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
                    ControlTipText ="Exit without saving  (ALT+C)"
                    UnicodeAccessKey =67
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
Dim strSQL As String
    If Left(Me.OpenArgs, 3) = "add" Then
        strSQL = "DELETE * FROM tblChart WHERE coaID = " & coaID
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
    If strSQL <> "" Then
        DoCmd.SetWarnings False
        FnLog (strSQL)
        DoCmd.RunSQL (strSQL)
        DoCmd.SetWarnings True
    End If

    If Forms.frmChart.xMyTreeview.Nodes.Count > 0 Then
        Forms!frmChart.PopulateTreeView (Forms.frmChart.xMyTreeview.SelectedItem.Key)
    End If
 End Sub

Private Sub cmdOK_Click()
Dim strSource As String
Dim tempcoaID As Long
Dim rs As Recordset
Dim strSQL As String

    On Error GoTo Err_cmdOK_Click
    
    If Me.coaName.Visible = False Then
        Me.coaName = Me.coaNameEng
    End If
    
    If Left(Me.OpenArgs, 3) = "add" Then
        strSQL = "SELECT tblChart.coaID, tblChart.coaName FROM tblChart WHERE tblChart.coaRef1 = " & coaRef1.Value & " AND tblChart.coaRef2 = " & coaRef2.Value & " AND tblChart.coaRef3 = " & coaRef3.Value & " AND tblChart.coaRef4 = " & coaRef4.Value & " AND tblChart.coaRef5 = " & coaRef5.Value

        Set rs = CurrentDb.OpenRecordset(strSQL)

        If rs.RecordCount > 0 Then
            MsgBox FnIsErr(3022) & Chr(10) & rs!coaName & " has this reference number.", vbExclamation
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
    tempcoaID = coaID
    
    If fsClosedYear And HasChanged Then
        If fsNextYearHasChanged Then
            MsgBox FnIsErr(7008), vbExclamation
        Else
            strSQL = "SELECT coaID FROM nexttblChart " & _
                     "WHERE coaRef1 = " & coaRef1 & " AND " & _
                     " coaRef2 = " & coaRef2 & " AND " & _
                     " coaRef3 = " & coaRef3 & " AND " & _
                     " coaRef4 = " & coaRef4 & " AND " & _
                    " coaRef5 = " & coaRef5

            Set rs = CurrentDb.OpenRecordset(strSQL)
            If rs.RecordCount > 0 Then
                strSQL = "UPDATE nexttblChart " & _
                        "SET coaName = '" & coaName & "'" & _
                        " WHERE coaRef1 = " & coaRef1 & " AND " & _
                        " coaRef2 = " & coaRef2 & " AND " & _
                        " coaRef3 = " & coaRef3 & " AND " & _
                        " coaRef4 = " & coaRef4 & " AND " & _
                        " coaRef5 = " & coaRef5
            Else
                strSQL = "INSERT INTO nexttblChart(coaID, lvlID, coaIDg, coaName, coaRef1, coaRef2, coaRef3, coaRef4, coaRef5) " & _
                        "VALUES(" & coaID & ", " _
                                  & lvlID & ", " _
                                  & coaIDg & ", '" _
                                  & coaName & "', " _
                                  & coaRef1 & ", " _
                                  & coaRef2 & ", " _
                                  & coaRef3 & ", " _
                                  & coaRef4 & ", " _
                                  & coaRef5 & ")"
            End If
            DoCmd.SetWarnings False
            FnLog (strSQL)
            DoCmd.RunSQL (strSQL)
            DoCmd.SetWarnings True
            rs.Close
            Set rs = Nothing
        End If
    End If

    DoCmd.Close acForm, "frmChartAdd"
    Forms!frmChart.PopulateTreeView ("a" & CStr(tempcoaID))
Exit_cmdOK_Click:
    Exit Sub
Err_cmdOK_Click:
   MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdOK_Click
End Sub


Private Sub coaName_AfterUpdate()
HasChanged = True
End Sub

Private Sub coaNameEng_AfterUpdate()
    If Len(Me.coaName) = 0 _
    Or Me.coaName.Visible = True Then
        Me.coaName = Me.coaNameEng
    End If
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
    coaNameEng.SetFocus
  Else
    Me.Filter = "coaID=" & Mid(Me.OpenArgs, 4, Len(Me.OpenArgs))
    Me.FilterOn = True
    lblcoaRef.Visible = False
    
    If DLookup("ctrTransl", "tblControl") = True Then
        Me.coaName.Visible = True
        Me.lblTransl.Visible = True
    Else
        Me.coaName.Visible = False
        Me.lblTransl.Visible = False
    End If
  End If
Exit_cmdAdd_Click:
    Exit Sub
Err_cmdAdd_Click:
   MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdAdd_Click

End Sub
