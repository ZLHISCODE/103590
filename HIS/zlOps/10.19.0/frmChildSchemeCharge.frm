VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmChildSchemeCharge 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2670
      Index           =   0
      Left            =   510
      ScaleHeight     =   2670
      ScaleWidth      =   6090
      TabIndex        =   0
      Top             =   630
      Width           =   6090
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   1
         Left            =   5730
         Picture         =   "frmChildSchemeCharge.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "��ѡ����ݼ���F3"
         Top             =   30
         Width           =   345
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   555
         TabIndex        =   2
         Top             =   900
         Width           =   1860
         _cx             =   3281
         _cy             =   2143
         Appearance      =   1
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
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         WordWrap        =   0   'False
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
End
Attribute VB_Name = "frmChildSchemeCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mlngKey As Long
Private mlngReferKey As Long
Private mfrmMain As Object
Private mbytMode As Byte
Private mblnAllowModify As Boolean
Private mblnDataChanged As Boolean
Private mblnReading As Boolean

Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Event AfterDataChanged()


Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    mlngKey = lngKey
    mbytMode = 2
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("�ؼ�״̬")
    
    If mlngKey > 0 Then
        If ExecuteCommand("��ȡ����") = False Then Exit Function
    End If

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function NewData(Optional ByVal lngReferKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 1
    
    mlngReferKey = lngReferKey
    If mlngReferKey > 0 Then
        mlngKey = mlngReferKey
        Call ExecuteCommand("��ȡ����")
        mlngKey = 0
    Else
        Call ExecuteCommand("�������")
    End If

    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = True
    
'    Call LocationObj(txt(2))
        
    NewData = True
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If lngLoop <> .Rows - 1 Then
                If Val(.RowData(lngLoop)) = 0 Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻����������������Ч���շ���Ŀ��"
                    .ShowCell .Row, .ColIndex("�շ���Ŀ")
                    LocationGrid vsf(0)
                    Exit Function
                End If
            End If
            
            If Val(.RowData(lngLoop)) > 0 Then
                If IsNumeric(.TextMatrix(lngLoop, .ColIndex("����"))) = False And Trim(.TextMatrix(lngLoop, .ColIndex("����"))) <> "" Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻�������շ���Ŀ����������Ϊ��ֵ�ͣ�"
                    .ShowCell .Row, .ColIndex("����")
                    LocationGrid vsf(0)
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("����"))) < 0 Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻���������������շ���Ŀ������[0-999999999]��"
                    .ShowCell .Row, .ColIndex("����")
                    LocationGrid vsf(0)
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("����"))) > 999999999 Then
                    ShowSimpleMsg "�� " & lngLoop & " ������̫�󣬱�������[0-999999999]��Χ�ڵ���ֵ��"
                    .ShowCell .Row, .ColIndex("����")
                    LocationGrid vsf(0)
                    Exit Function
                End If
                
            End If
        Next
    End With
    
    ValidData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset, ByRef lngKey As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lngLoop As Long
    Dim strTmp As String

    On Error GoTo errHand

    strSQL = "ZL_�������Ѳο�_DELETE(" & lngKey & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If Val(.RowData(lngLoop)) > 0 And Val(.TextMatrix(lngLoop, .ColIndex("����"))) >= 0 Then
                strSQL = "ZL_�������Ѳο�_INSERT(" & lngKey & "," & Val(.RowData(lngLoop)) & "," & Val(.TextMatrix(lngLoop, .ColIndex("����"))) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    SaveData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            If mblnAllowModify Then
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
            Else
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            End If
            Call .AppendColumn("�շ���Ŀ", 2700, flexAlignLeftCenter, flexDTString, "", "����", True)
            Call .AppendColumn("���", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���㵥λ", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 600, flexAlignLeftCenter, flexDTDecimal, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            
            If mblnAllowModify Then
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("�շ���Ŀ"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditText)
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
            End If
            
            .AppendRows = True
        End With
        cmd(1).Enabled = mblnAllowModify
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
        
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 And mbytMode = 2 Then blnAllowModify = False
        mclsVsf.AllowEdit = blnAllowModify
        cmd(1).Enabled = blnAllowModify
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
    
        mclsVsf.ClearGrid
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
        
        mclsVsf.ClearGrid
        
        gstrSQL = "SELECT D.ID,b.���� As ���,d.���㵥λ,D.����,A.���� " & _
                "FROM �������Ѳο� A,�շ���ĿĿ¼ D,�շ���Ŀ��� B " & _
                "WHERE A.ϸĿid=D.ID And B.����=d.��� And A.����ID=[1]"
                
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)
        
    End Select

    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim lngLoop As Long
    
    Select Case Index
    Case 1

        gstrSQL = GetPublicSQL(SQL.������Ŀѡ��)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)

        If ShowPubSelect(Me, cmd(1), 3, "����,1200,0,;����,2700,0,;���,900,0,;��λ,900,0,", Me.Name & "\������Ŀ��ѡ", "����±���ѡ��һ������������Ŀ", rsData, rs, 8790, 4800, True) = 1 Then
            With vsf(0)
                For lngLoop = 1 To rs.RecordCount
                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) = False Then

                        If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1

                        .TextMatrix(.Rows - 1, mclsVsf.ColIndex("�շ���Ŀ")) = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(.Rows - 1, mclsVsf.ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                        .TextMatrix(.Rows - 1, mclsVsf.ColIndex("���㵥λ")) = zlCommFun.NVL(rs("��λ").Value)
                        
                        If Val(.TextMatrix(.Row, .ColIndex("����"))) = 0 Then
                            .TextMatrix(.Rows - 1, mclsVsf.ColIndex("����")) = "1"
                        Else
                            .TextMatrix(.Rows - 1, mclsVsf.ColIndex("����")) = .TextMatrix(.Row, .ColIndex("����"))
                        End If
                        
                        .RowData(.Rows - 1) = zlCommFun.NVL(rs("ID").Value, 0)

                        DataChanged = True
                    End If

                    rs.MoveNext
                Next
            End With
        End If
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        KeyCode = 0
        If cmd(0).Enabled And cmd(0).Visible Then
            Call cmd_Click(0)
        End If
    End If
End Sub

Private Sub Form_Load()
'    picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub mclsVsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width - cmd(1).Width - 30, picPane(Index).Height
        cmd(1).Move vsf(0).Left + vsf(0).Width + 15
        mclsVsf.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    '�༭����
    Call mclsVsf.AfterEdit(Row, Col)
    
    DataChanged = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf(0).RowData(Row)) <= 0 Or Val(vsf(0).TextMatrix(Row, vsf(0).ColIndex("����"))) < 0)
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset

    With vsf(0)
        If Col = .ColIndex("�շ���Ŀ") Then

            gstrSQL = GetPublicSQL(SQL.������Ŀѡ��)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
            If ShowPubSelect(Me, vsf(0), 3, "����,1200,0,0;����,3000,0,0;���,900,0,0;��λ,900,0,0", Me.Name & "\������Ŀѡ��", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then
                If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
    
                .EditText = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("�շ���Ŀ")) = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("���㵥λ")) = zlCommFun.NVL(rs("��λ").Value)
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                DataChanged = True
                Call mclsVsf.LocationNextCell
            End If
            
            Call mclsVsf.SetFocus(, , True)
        End If
    End With
End Sub

Private Sub vsf_ChangeEdit(Index As Integer)
    With vsf(Index)
        Select Case .Col
        Case .ColIndex("����")
            .TextMatrix(.Row, .Col) = .EditText
        End Select

    End With
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    
    With vsf(0)
        If KeyCode = vbKeyReturn Then
            If Col = .ColIndex("�շ���Ŀ") Then
                
                If InStr(vsf(0).EditText, "'") > 0 Then
                    KeyCode = 0
                    vsf(0).EditText = ""
                    Exit Sub
                End If

                strText = UCase(vsf(0).EditText)
                bytMode = GetApplyMode(strText)
                
                gstrSQL = GetPublicSQL(SQL.������Ŀ����, bytMode)
                strText = strText & "%"
                If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                    strTmp = strText
                Else
                    strTmp = "%" & strText
                End If
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)

                If ShowPubSelect(Me, vsf(0), 2, "����,1200,0,0;����,3000,0,0;���,900,0,0;��λ,900,0,0", Me.Name & "\������Ŀ����", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                        Exit Sub
                    End If

                    .EditText = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("�շ���Ŀ")) = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("���㵥λ")) = zlCommFun.NVL(rs("��λ").Value)
                    .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                    
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

                    DataChanged = True
                    Call mclsVsf.LocationNextCell
                Else
                    KeyCode = 0

                    .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                    .EditText = .Cell(flexcpData, Row, Col)
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)

                End If
                
                Call mclsVsf.SetFocus(, , True)
            Else
                Call mclsVsf.LocationNextCell
            End If
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    '�༭����
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub



