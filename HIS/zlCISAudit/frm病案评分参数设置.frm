VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm�������ֲ������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������ֲ�������"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "frm�������ֲ�������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   345
      TabIndex        =   4
      Top             =   5490
      Width           =   1100
   End
   Begin VB.Frame fraStatus 
      Caption         =   "���ֿ��ҷ�Χ"
      Height          =   5250
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   6390
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   4800
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   270
         Width           =   6255
         _cx             =   11033
         _cy             =   8467
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5190
      TabIndex        =   1
      Top             =   5490
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3960
      TabIndex        =   0
      Top             =   5475
      Width           =   1100
   End
End
Attribute VB_Name = "frm�������ֲ�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mblnOK As Boolean
Private mfrmMain As Object
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mclsVsfNo As clsVsf

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim intCol As Integer
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim varTmp As Variant
    Dim varAry As Variant
    Dim blnAllowModify As Boolean

    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "��ʼ����"
            
            Set mclsVsf = New clsVsf
            With mclsVsf
                Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
                Call .ClearColumn
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)

                Call .AppendColumn("����", 1590, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("����", 3690, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
                
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditCommand)
                .IndicatorCol = 0
                Set .IndicatorIcon = GetImageList(16).ListImages("��ǰ").Picture
          
                .AppendRows = True
            End With
             
        '--------------------------------------------------------------------------------------------------------------
        Case "��ȡ����"
            
            On Error Resume Next
              
            
            strTmp = Trim(zlDatabase.GetPara("���ֿ��ҷ�Χ", ParamInfo.ϵͳ��, 1562, "", Array(vsf(0)), True))
            gstrSQL = "Select ID,���,���� From ��Ա��"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            
            gstrSQL = "Select a.ID,a.����,a.���� From ���ű� a,��������˵�� b Where a.ID=b.����id And b.��������='�ٴ�' and ( TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null)"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            
            With vsf(0)
                .Rows = 2
                varTmp = Split(strTmp, ";")
                For intCount = 0 To UBound(varTmp)
                    varAry = Split(varTmp(intCount), ",")
                    rs.Filter = ""
                    rs.Filter = "ID=" & Val(varAry(0))
                    If rs.RecordCount > 0 Then
                        
                        If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, .ColIndex("����")) = AppendCode(rs("����").Value, rs("���").Value)
                        .RowData(.Rows - 1) = rs("ID").Value
                        
                        For intCol = 1 To UBound(varAry)
                            rsTmp.Filter = ""
                            rsTmp.Filter = "ID=" & Val(varAry(intCol))
                            If rsTmp.RecordCount > 0 Then
                                If .TextMatrix(.Rows - 1, .ColIndex("����")) = "" Then
                                    .TextMatrix(.Rows - 1, .ColIndex("����")) = AppendCode(rsTmp("����").Value, rsTmp("����").Value)
                                    .TextMatrix(.Rows - 1, .ColIndex("����id")) = rsTmp("ID").Value
                                Else
                                    .TextMatrix(.Rows - 1, .ColIndex("����")) = .TextMatrix(.Rows - 1, .ColIndex("����")) & vbCrLf & AppendCode(rsTmp("����").Value, rsTmp("����").Value)
                                    .TextMatrix(.Rows - 1, .ColIndex("����id")) = .TextMatrix(.Rows - 1, .ColIndex("����id")) & "," & rsTmp("ID").Value
                                End If
                            End If
                        Next
                    End If
                Next
                .AutoSize .ColIndex("����"), .ColIndex("����")
            End With
            
        '--------------------------------------------------------------------------------------------------------------
        Case "У������"
            
        '--------------------------------------------------------------------------------------------------------------
        Case "��������"
            
            
            
            
            
        End Select
    Next

    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    cmdOK.Tag = IIf(blnData, "Changed", "")
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = (cmdOK.Tag = "Changed")
End Property

'######################################################################################################################

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strTmp      As String
    Dim intCount    As Long
    strTmp = ""
    With vsf(0)
        For intCount = 1 To .Rows - 1
            If Val(.RowData(intCount)) > 0 Then
                strTmp = strTmp & ";" & Val(.RowData(intCount)) & "," & Trim(.TextMatrix(intCount, .ColIndex("����id")))
            End If
        Next
    End With
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    If Len(strTmp) > 2000 Then
        ShowSimpleMsg "������Ȩ��̫�࣬�����˲���ֵ�����洢��Χ��"
        Exit Sub
    End If
    Call SetPara("���ֿ��ҷ�Χ", strTmp, "1562")
    
    Unload Me

End Sub

Private Sub Form_Load()
    mblnOK = False
   
    If ExecuteCommand("��ʼ����") = False Then Exit Sub
    If ExecuteCommand("��ȡ����") = False Then Exit Sub
    
    vsf(0).Refresh
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    Set mclsVsf = Nothing
    Set mclsVsfNo = Nothing
    
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    
    With vsf(0)
        Cancel = Not (Val(.RowData(Row)) > 0 And Trim(.TextMatrix(Row, .ColIndex("����id"))) <> "")
        If Cancel = False Then DataChanged = True
    End With
    
End Sub


Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    '�༭����
    Select Case Index
        Case 0
            Call mclsVsf.AfterEdit(Row, Col)
    End Select
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '�༭����
    Select Case Index
        Case 0
            Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    End Select
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Select Case Index
    Case 0
        mclsVsf.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Select Case Index
    Case 0
        mclsVsf.AppendRows = True
    End Select
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim varTmp As Variant
    Dim bytRet As Byte
    Dim strTmp As String
    Dim strTmpID As String
    Dim intCount As Integer
    
    With vsf(Index)
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
            Select Case Col
            '----------------------------------------------------------------------------------------------------------
            Case .ColIndex("����")
                
                Set rsData = gclsPackage.GetOperationPerson
                bytRet = ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,", Me.Name & "\�����Աѡ��", "����±���ѡ��һ�������Ա", rsData, rs, 8790, 4500, False, Val(.RowData(Row)))
                            
                If bytRet = 1 Then
                    
                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value), False) = False Then
                        
                        .EditText = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .TextMatrix(Row, .ColIndex("����")) = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        DataChanged = True
                    End If

                    
                    mclsVsf.AppendRows = True
        
                End If
            '----------------------------------------------------------------------------------------------------------
            Case .ColIndex("����")
            
                Set rs = gclsPackage.GetDeptSelect
                Set rsData = CopyRecordStruct(rs)
                Call CopyRecordData(rs, rsData)
                
                If .TextMatrix(Row, .ColIndex("����id")) <> "" Then
                    varTmp = Split(.TextMatrix(Row, .ColIndex("����id")), ",")
                    For intCount = 0 To UBound(varTmp)
                        rsData.Filter = ""
                        rsData.Filter = "ID=" & Val(varTmp(intCount))
                        If rsData.RecordCount > 0 Then
                            rsData("ѡ��").Value = 1
                        End If
                    Next
                End If
                rsData.Filter = ""
                If rsData.RecordCount > 0 Then rsData.MoveFirst
                
                bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,1200,0,;����,900,0,", Me.Name & "\���˿���ѡ��", "����±���ѡ��һ���������˿���", rsData, rs, 8790, 4500, True)
                            
                If bytRet = 1 Then
                    
                    If rs.RecordCount > 0 Then rs.MoveFirst
                    strTmp = ""
                    strTmpID = ""
                    Do While Not rs.EOF
                        strTmp = strTmp & vbCrLf & AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("����").Value))
                        strTmpID = strTmpID & "," & zlCommFun.NVL(rs("ID").Value, 0)
                        rs.MoveNext
                    Loop
                    If strTmp <> "" Then strTmp = Mid(strTmp, 3)
                    If strTmpID <> "" Then strTmpID = Mid(strTmpID, 2)
                    
                    .EditText = strTmp
                    .TextMatrix(Row, .ColIndex("����")) = strTmp
                    .TextMatrix(Row, .ColIndex("����id")) = strTmpID
                    
                    DataChanged = True

                    .AutoSize .ColIndex("����"), .ColIndex("����")
                    mclsVsf.AppendRows = True
        
                End If
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case 1
            
        End Select
    End With
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsf.KeyDown(KeyCode, Shift)
    End Select
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim StrText As String
    Dim bytRet As Byte
    
    With vsf(Index)
        
        If InStr(.EditText, "'") > 0 Then
            KeyCode = 0
            .EditText = ""
            Exit Sub
        End If
                            
        StrText = .EditText
        
        Select Case Index
        '----------------------------------------------------------------------------------------------------------
        Case 0
            If KeyCode = vbKeyReturn Then
                If Col = .ColIndex("����") Then

                    Set rsData = gclsPackage.GetOperationPerson(UCase(StrText))
                    
                    If ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,", Me.Name & "\�����Ա����", "����±���ѡ��һ�������Ա", rsData, rs, 8790, 4500, , Val(.RowData(Row)), , True) = 1 Then
    
                        If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ա��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                               
                        .EditText = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .Cell(flexcpData, Row, Col) = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .TextMatrix(Row, .ColIndex("����")) = AppendCode(zlCommFun.NVL(rs("����").Value), zlCommFun.NVL(rs("���").Value))
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        DataChanged = True
                    Else
                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                        .EditText = .Cell(flexcpData, Row, Col)
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    End If
    
                End If
            Else
                DataChanged = True
            End If

        End Select
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    
    '�༭����,������
    Select Case Index
    Case 0
        Call mclsVsf.KeyPress(KeyAscii)
    End Select
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsf.KeyPressEdit(KeyAscii)
    End Select
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Select Case Index
            Case 0
                Call mclsVsf.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        End Select
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsf.EditSelAll
    End Select
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Select Case Index
        Case 0
            Call mclsVsf.BeforeEdit(Row, Col, Cancel)
    End Select
End Sub
