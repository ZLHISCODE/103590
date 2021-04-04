VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Begin VB.Form frmChildStationPerson 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   810
      ScaleHeight     =   3900
      ScaleWidth      =   6180
      TabIndex        =   0
      Top             =   1095
      Width           =   6180
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2145
         Index           =   0
         Left            =   690
         TabIndex        =   1
         Top             =   1050
         Width           =   3990
         _cx             =   7038
         _cy             =   3784
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
         TreeColor       =   -2147483632
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
      Begin MSComctlLib.TabStrip tbs 
         Height          =   300
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   529
         MultiRow        =   -1  'True
         Style           =   2
         TabMinWidth     =   529
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "2"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmChildStationPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'���������弶��������

Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mfrmMain As Object
Private mblnAllowModify As Boolean
Private mlngKey As Long
Private mlngDeptKey As Long

Private mobjStateInfo As CommandBarControl
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Event AfterDataChanged()
Public Event AfterMakeCharge()

'######################################################################################################################
'�������Զ�����̻���

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    
    
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
    
    tbs.Enabled = Not mblnDataChanged
    
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
    
    If ExecuteCommand("��ʼ�ؼ�") = False Then Exit Function
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long, _
                            Optional ByVal lngDeptKey As Long, _
                            Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mlngDeptKey = lngDeptKey
    mblnAllowModify = blnAllowModify
    mlngKey = lngKey
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("�ؼ�״̬")
    
    If ExecuteCommand("��ȡ�ڼ�") = False Then Exit Function
    If ExecuteCommand("��ȡ����") = False Then Exit Function

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    
'    With vsf(0)
'        For lngLoop = 1 To .Rows - 1
'            If Val(.RowData(lngLoop)) > 0 And InStr(.TextMatrix(lngLoop, .ColIndex("��λ")), "����ҽ��") > 0 Then
'                Exit For
'            End If
'        Next
'        If lngLoop = .Rows Then
'            ShowSimpleMsg " ����ָ������������ҽ����"
''            Call LocationGrid(vsf(4), 1, .ColIndex("����"))
'            Exit Function
'        End If
'    End With
    
    ValidData = True
    
End Function

Public Function ClearData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Call ExecuteCommand("�������")
    
    ClearData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lngLoop As Long
    Dim lngRow As Long
    
    On Error GoTo errHand

    strSQL = "zl_����������Ա_Delete(" & mlngKey & "," & Val(tbs.SelectedItem.Tag) & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(0)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 Then
                strSQL = "zl_����������Ա_Insert(" & mlngKey & ",'" & .TextMatrix(lngRow, .ColIndex("��λ")) & "'," & Val(.RowData(lngRow)) & ",'" & .TextMatrix(lngRow, .ColIndex("����")) & "'," & Val(tbs.SelectedItem.Tag) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With

    SaveData = True
    
    Exit Function
    
    '
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    Dim cbrCustom As CommandBarControlCustom
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)
    
    cbsMain.Options.LargeIcons = False
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������

    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewParent, "�����ڼ�", , , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_DeleteParent, "ɾ���ڼ�", , , xtpButtonIconAndCaption)
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "������Ա", True, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ����Ա", , , xtpButtonIconAndCaption)


    Set mobjStateInfo = NewToolBar(objBar, xtpControlLabel, 0, "")
    mobjStateInfo.Flags = xtpFlagRightAlign
        
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim objArray As Variant
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        '������Ա
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("��λ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���", 900, flexAlignLeftCenter, flexDTString, "", "", True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        Call InitCommandBar
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
    
        mclsVsf.ClearGrid
        mobjStateInfo.Caption = " "
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 Then blnAllowModify = False
        
        With mclsVsf

            If blnAllowModify Then


                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                Call .InitializeEdit(True, True, True)

                '������λ
                '------------------------------------------------------------------------------------------------------
                gstrSQL = "SELECT ����||'-'||���� As ���� FROM ������λ Order by ����"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                Call .InitializeEditColumn(.ColIndex("��λ"), True, vbVsfEditCombox, vsf(0).BuildComboList(rs, "����", "����"))
                
                Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditCommand)
                
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture

            Else
                Call .InitializeEdit(False, False, False)
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            End If
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ�ڼ�"
        
        
        intLoop = 0
        
        tbs.Tabs.Clear
        gstrSQL = "Select a.�ڼ� From ����������Ա a Where a.��¼id=[1] Group By a.�ڼ�"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            Do While Not rs.EOF
                intLoop = intLoop + 1
                tbs.Tabs.Add intLoop, , CStr(intLoop)
                tbs.Tabs(intLoop).Tag = rs("�ڼ�").Value
                rs.MoveNext
            Loop
        Else
            tbs.Tabs.Add 1, , "1"
            tbs.Tabs(1).Tag = 1
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
            
        mclsVsf.ClearGrid
        mobjStateInfo.Caption = " "
        
        gstrSQL = "Select A.��Աid As ID,a.��λ,B.���,a.���� From ����������Ա a,��Ա�� b,������λ c Where c.����=a.��λ And a.��¼id=[1] And a.�ڼ�=[2] And a.��Աid=b.ID(+) order by c.����"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, Val(tbs.SelectedItem.Tag))
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)
        
        cbsMain.RecalcLayout
        
    End Select
    
    ExecuteCommand = True
    
    Exit Function
    
    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Public Property Get Body(ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intRow As Integer
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewParent                         '�����ڼ�
        
        intRow = tbs.Tabs.Count + 1
        tbs.Tabs.Add intRow, , CStr(intRow)
        tbs.Tabs(intRow).Tag = intRow
        
        tbs.Tabs(intRow).Selected = True
        
        Call ExecuteCommand("��ȡ����")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_DeleteParent                      'ɾ���ڼ�
        
        tbs.Tabs.Remove tbs.SelectedItem.Index
        DataChanged = True
        
        Call ExecuteCommand("��ȡ����")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem                           '������Ա
        
        Call mclsVsf.AppendRow
        Call mclsVsf.SetFocus(, 1)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                            'ɾ����Ա
        
        Call mclsVsf.DeleteRow(vsf(0).Row)

    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl

    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_MakeCharge

        With CommandBar.Controls

            .DeleteAll
            
            
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 1, "�շѵ���(&1)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 2, "���ʵ���(&2)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 3, "��ķ���(&3)")
            With cbsMain.KeyBindings
                .Add FCONTROL, vbKeyN, conMenu_Edit_MakeCharge * 2 + 1
                .Add FCONTROL, vbKeyB, conMenu_Edit_MakeCharge * 2 + 2
            End With
            
        End With
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '���������ؼ�Resize����
    picPane.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand
    
    blnAllowModify = mblnAllowModify
    If mlngKey = 0 Then blnAllowModify = False
        
    With vsf(0)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewParent
        
            Control.Enabled = blnAllowModify And Control.Visible And DataChanged = False
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_DeleteParent
            
            Control.Enabled = blnAllowModify And Control.Visible And DataChanged = False And Val(tbs.SelectedItem.Tag) > 1
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem
            
    '        Control.Visible = IsPrivs(mstrPrivs, "��ҩ׼��") Or IsPrivs(mstrPrivs, "��ҩ�Ǽ�")
            Control.Enabled = blnAllowModify And Control.Visible
                        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            
    '        Control.Visible = IsPrivs(mstrPrivs, "��ҩ׼��") Or IsPrivs(mstrPrivs, "��ҩ�Ǽ�")
            
            Control.Enabled = blnAllowModify And Control.Visible And (.TextMatrix(.Row, .ColIndex("��λ")) <> "" Or .TextMatrix(.Row, .ColIndex("����")) <> "")
            
            
        End Select
    End With
errHand:

End Sub

'######################################################################################################################
'���������弰��ؼ����¼�����

Private Sub Form_Resize()
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjStateInfo = Nothing
    Set mclsVsf = Nothing
End Sub

Private Sub mclsVsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(0)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub picPane_Resize()
    On Error Resume Next
    
    tbs.Move 30, 30, picPane.Width - 30, tbs.Height
    vsf(0).Move 0, tbs.Top + tbs.Height + 30, picPane.Width, picPane.Height - (tbs.Top + tbs.Height + 30)
    mclsVsf.AppendRows = True
    
End Sub

Private Sub tbs_Click()
    Call ExecuteCommand("��ȡ����")
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    '�༭����
    With vsf(Index)

        Call mclsVsf.AfterEdit(Row, Col)
                
        Select Case Col
        Case .ColIndex("��λ")
            If .ComboIndex > -1 Then
                .TextMatrix(Row, Col) = zlCommFun.GetNeedName(.ComboItem(.ComboIndex))
            End If
        End Select

    End With
    
    DataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)

    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)

End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)

    mclsVsf.AppendRows = True

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
    Dim bytRet As Byte
    Dim strTmp As String
    
    With vsf(Index)
        If Col = .ColIndex("����") Then
                
            strTmp = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("��λ")))
            
            gstrSQL = "Select �Ƿ�Ψһ,�Ƿ�ҽ��,�Ƿ�ʿ From ������λ Where ����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTmp)
            If rs.BOF = False Then
                If zlCommFun.NVL(rs("�Ƿ�ҽ��").Value, 0) = 1 Then strTmp = "ҽ��"
                If zlCommFun.NVL(rs("�Ƿ�ʿ").Value, 0) = 1 Then strTmp = "��ʿ"
            Else
                strTmp = "ҽ��"
            End If
            
            gstrSQL = GetPublicSQL(SQL.��Ա��Ϣѡ��)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strTmp, mlngDeptKey)

            If ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1500,0,;����,900,0,;����,1200,0,", Me.Name & "\��Ա��Ϣѡ��", "����±���ѡ��һ����Ա", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ա��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
                       
                .EditText = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                
                DataChanged = True
            End If
        End If
    End With
End Sub

Private Sub vsf_ComboDropDown(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    
    With vsf(Index)

            Select Case Col
            Case .ColIndex("��λ")
                
                Call mclsVsf.ComboLocation(Row, Col)

            End Select

    End With
End Sub

Private Sub vsf_DblClick(Index As Integer)

    Call mclsVsf.DbClick

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
    Dim bytRet As Byte
    Dim strClass As String
    
    With vsf(Index)
        If KeyCode = vbKeyReturn Then
        
            If InStr(.EditText, "'") > 0 Then
                KeyCode = 0
                .EditText = ""
                Exit Sub
            End If
            strText = UCase(.EditText)
            bytMode = GetApplyMode(strText)
            strText = strText & "%"
            strTmp = IIf(ParamInfo.��Ŀ����ƥ�䷽ʽ = 1, strText, "%" & strText)
                    
                
            If Col = .ColIndex("����") Then
                
                strClass = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("��λ")))
                
                gstrSQL = "Select �Ƿ�Ψһ,�Ƿ�ҽ��,�Ƿ�ʿ From ������λ Where ����=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strClass)
                strClass = ""
                If rs.BOF = False Then
                    If zlCommFun.NVL(rs("�Ƿ�ҽ��").Value, 0) = 1 Then strClass = "ҽ��"
                    If zlCommFun.NVL(rs("�Ƿ�ʿ").Value, 0) = 1 Then strClass = "��ʿ"
                Else
                    strClass = "ҽ��"
                End If
                
                gstrSQL = GetPublicSQL(SQL.��Ա��Ϣ����, bytMode)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strClass, mlngDeptKey, strText, strTmp)
    
                If ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1500,0,;����,900,0,;����,1200,0,", Me.Name & "\��Ա��Ϣ����", "����±���ѡ��һ����Ա", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "ѡ�����Ա��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                        Exit Sub
                    End If
                           
                    .EditText = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    DataChanged = True
                Else
                    .Cell(flexcpData, Row, Col) = .EditText
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    DataChanged = True
                End If
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

