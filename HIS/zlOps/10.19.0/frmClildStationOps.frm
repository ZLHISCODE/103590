VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~4.OCX"
Begin VB.Form frmClildStationOps 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2985
      Index           =   1
      Left            =   4320
      ScaleHeight     =   2985
      ScaleWidth      =   4290
      TabIndex        =   2
      Top             =   2670
      Width           =   4290
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2145
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   285
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
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2985
      Index           =   0
      Left            =   840
      ScaleHeight     =   2985
      ScaleWidth      =   4290
      TabIndex        =   0
      Top             =   870
      Width           =   4290
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2145
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   285
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmClildStationOps.frx":0000
      Left            =   1050
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmClildStationOps"
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
Private mstrPrivs As String
Private mobjStateInfo As CommandBarControl

Private WithEvents mclsVsfBefore As clsVsf
Attribute mclsVsfBefore.VB_VarHelpID = -1
Private WithEvents mclsVsfAfter As clsVsf
Attribute mclsVsfAfter.VB_VarHelpID = -1

Public Event AfterDataChanged()
Public Event AfterMakeCharge()

'######################################################################################################################
'�������Զ�����̻���

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
    
    If ExecuteCommand("��ʼ�ؼ�") = False Then Exit Function
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngKey = lngKey
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("�ؼ�״̬")
    
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

    '��������
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_�����������_DELETE(" & mlngKey & ",1)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(0)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 Or .TextMatrix(lngRow, .ColIndex("��������")) <> "" Then

                If Left(.TextMatrix(lngRow, .ColIndex("���뷽ʽ")), 1) = 1 Then
                    strSQL = "zl_�����������_Insert(" & mlngKey & ",1," & Abs(Val(.TextMatrix(lngRow, .ColIndex("��Ҫ����")))) & ",'" & .TextMatrix(lngRow, .ColIndex("��������")) & "',Null," & Val(.RowData(lngRow)) & ")"
                Else
                    strSQL = "zl_�����������_Insert(" & mlngKey & ",1," & Abs(Val(.TextMatrix(lngRow, .ColIndex("��Ҫ����")))) & ",'" & .TextMatrix(lngRow, .ColIndex("��������")) & "'," & Val(.RowData(lngRow)) & ",Null)"
                End If
                
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    '��������
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_�����������_DELETE(" & mlngKey & ",2)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(1)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 Or .TextMatrix(lngRow, .ColIndex("��������")) <> "" Then

                If Left(.TextMatrix(lngRow, .ColIndex("���뷽ʽ")), 1) = 1 Then
                    strSQL = "zl_�����������_Insert(" & mlngKey & ",2," & Abs(Val(.TextMatrix(lngRow, .ColIndex("��Ҫ����")))) & ",'" & .TextMatrix(lngRow, .ColIndex("��������")) & "',Null," & Val(.RowData(lngRow)) & ")"
                Else
                    strSQL = "zl_�����������_Insert(" & mlngKey & ",2," & Abs(Val(.TextMatrix(lngRow, .ColIndex("��Ҫ����")))) & ",'" & .TextMatrix(lngRow, .ColIndex("��������")) & "'," & Val(.RowData(lngRow)) & ",Null)"
                End If
                
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
    
    Set objPopup = NewToolBar(objBar, xtpControlPopup, conMenu_Edit_MakeCharge, "�������������ɷ���", True, , xtpButtonIconAndCaption)

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
        Set mclsVsfBefore = New clsVsf
        With mclsVsfBefore
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("���뷽ʽ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������", 2400, flexAlignLeftCenter, flexDTString, "", "��������", True)
            Call .AppendColumn("��Ҫ����", 810, flexAlignCenterCenter, flexDTBoolean, "", "ȱʡ", True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        Set mclsVsfAfter = New clsVsf
        With mclsVsfAfter
            Call .Initialize(Me.Controls, vsf(1), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("���뷽ʽ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������", 2400, flexAlignLeftCenter, flexDTString, "", "��������", True)
            Call .AppendColumn("��Ҫ����", 810, flexAlignCenterCenter, flexDTBoolean, "", "ȱʡ", True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        Call InitCommandBar
        
        Dim objPane As Pane
    
        Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing)
        objPane.Title = "��������"
        objPane.Options = PaneNoCaption
    
        Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, objPane)
        objPane.Title = "��������"
        objPane.Options = PaneNoCaption
    
        dkpMain.SetCommandBars cbsMain
        dkpMain.Options.ThemedFloatingFrames = True
        dkpMain.Options.UseSplitterTracker = False
        dkpMain.Options.AlphaDockingContext = True
        dkpMain.Options.HideClient = True
    
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
    
        mclsVsfBefore.ClearGrid
        mclsVsfAfter.ClearGrid
        mobjStateInfo.Caption = " "
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 Then blnAllowModify = False
        
        With mclsVsfBefore
            If blnAllowModify Then
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("���뷽ʽ"), True, vbVsfEditCombox, "1-����|2-����")
                Call .InitializeEditColumn(.ColIndex("��������"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("��Ҫ����"), True, vbVsfEditCheck)
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
            Else
                Call .InitializeEdit(False, False, False)
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            End If
        End With
    
        With mclsVsfAfter
            If blnAllowModify Then
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("���뷽ʽ"), True, vbVsfEditCombox, "1-����|2-����")
                Call .InitializeEditColumn(.ColIndex("��������"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("��Ҫ����"), True, vbVsfEditCheck)
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
            Else
                Call .InitializeEdit(False, False, False)
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��״̬"
    
        mobjStateInfo.Caption = " "
        
        gstrSQL = "Select No,��¼���� From ������������ Where ��¼id=[1] And ��������=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey, 4)
        If rs.BOF = False Then
            If IsNull(rs("No").Value) = False Then
                Select Case rs("��¼����").Value
                Case 1
                    mobjStateInfo.Caption = "�������շѵ������ݺţ�" & rs("No").Value
                Case 2
                    mobjStateInfo.Caption = "�����ɼ��ʵ������ݺţ�" & rs("No").Value
                End Select
            End If
        End If
        cbsMain.RecalcLayout
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"

        '��������
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfBefore.ClearGrid
        gstrSQL = GetPublicSQL(SQL.�����������)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 1)
        If rs.BOF = False Then Call mclsVsfBefore.LoadGrid(rs)
        
        '��������
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfAfter.ClearGrid
        gstrSQL = GetPublicSQL(SQL.�����������)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 2)
        If rs.BOF = False Then Call mclsVsfAfter.LoadGrid(rs)
        
        Call ExecuteCommand("ˢ��״̬")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "���������շѵ�"
    
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("��������ɣ�����Զ�ɾ�������ϣ�ȷ��Ҫ����������������Ϊ�շѵ����µ�����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        gstrSQL = "Select ҽ��id From ����������¼ Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then
            strTmp = MakeChargeBill(Val(rs("ҽ��id").Value), 1, "����", False, mstrPrivs)
            If strTmp <> "" Then
                ShowSimpleMsg "�����շѵ��Ѿ����ɣ����ݺţ�" & strTmp
                Call ExecuteCommand("ˢ��״̬")
                RaiseEvent AfterMakeCharge
            End If
        End If
        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�����������ʵ�"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("��������ɣ�����Զ�ɾ�������ϣ�ȷ��Ҫ����������������Ϊ���ʵ����µ�����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        gstrSQL = "Select ҽ��id From ����������¼ Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then
            strTmp = MakeChargeBill(Val(rs("ҽ��id").Value), 2, "����", False, mstrPrivs)
            If strTmp <> "" Then
                ShowSimpleMsg "�������ʵ��Ѿ����ɣ����ݺţ�" & strTmp
                Call ExecuteCommand("ˢ��״̬")
                RaiseEvent AfterMakeCharge
            End If
        End If
        
        Exit Function
    '------------------------------------------------------------------------------------------------------------------
    Case "����������ѵ�"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("��������ɣ�����Զ�ɾ�������ϣ�ȷ��Ҫ����������������Ϊ��ѵ����µ�����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        gstrSQL = "Select ҽ��id From ����������¼ Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then
        
            strTmp = MakeChargeBill(Val(rs("ҽ��id").Value), 2, "����", True, mstrPrivs)
            If strTmp <> "" Then
                ShowSimpleMsg "������ѵ��Ѿ����ɣ����ݺţ�" & strTmp
                Call ExecuteCommand("ˢ��״̬")
                RaiseEvent AfterMakeCharge
            End If
        
        End If
        
        Exit Function
        
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

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intRow As Integer
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 1

        Call ExecuteCommand("���������շѵ�")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 2

        Call ExecuteCommand("�����������ʵ�")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 3
                
        Call ExecuteCommand("����������ѵ�")
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

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand
    
    blnAllowModify = mblnAllowModify
    If mlngKey = 0 Then blnAllowModify = False

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge
        
        Control.Enabled = blnAllowModify And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2 + 1
        
'        Control.Visible = (mstr������Դ = "����" And IsPrivs(mstrPrivs, "���ɸ���"))
        Control.Enabled = blnAllowModify And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2 + 2, conMenu_Edit_MakeCharge * 2 + 3

        Control.Enabled = blnAllowModify And Control.Visible
        
    End Select

errHand:

End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

'######################################################################################################################
'���������弰��ؼ����¼�����

Private Sub Form_Unload(Cancel As Integer)
    Set mobjStateInfo = Nothing
    Set mclsVsfBefore = Nothing
    Set mclsVsfAfter = Nothing
End Sub

Private Sub mclsVsfAfter_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf(1).RowData(Row)) > 0 Then
        DataChanged = True
    End If
End Sub

Private Sub mclsVsfAfter_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(1)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub mclsVsfBefore_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf(0).RowData(Row)) > 0 Then
        DataChanged = True
    End If
End Sub

Private Sub mclsVsfBefore_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(0)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsfBefore.AppendRows = True
    Case 1
        vsf(1).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsfAfter.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    '�༭����
    Select Case Index
    Case 0
        Call mclsVsfBefore.AfterEdit(Row, Col)
    Case 1
        Call mclsVsfAfter.AfterEdit(Row, Col)
    End Select

    DataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Select Case Index
    Case 0
        Call mclsVsfBefore.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    Case 1
        Call mclsVsfAfter.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    End Select
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Select Case Index
    Case 0
        mclsVsfBefore.AppendRows = True
    Case 1
        mclsVsfAfter.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Select Case Index
    Case 0
        mclsVsfBefore.AppendRows = True
    Case 1
        mclsVsfAfter.AppendRows = True
    End Select
End Sub


Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Select Case Index
    Case 0
        mclsVsfBefore.AppendRows = True
    Case 1
        mclsVsfAfter.AppendRows = True
    End Select
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    Dim strTmp As String
    
    With vsf(Index)
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0                               '��������
            
            If Col = .ColIndex("��������") Then

                If Val(Left(.TextMatrix(Row, .ColIndex("���뷽ʽ")), 1)) = 1 Then
                    '���Ʊ���
                    gstrSQL = GetPublicSQL(SQL.������Ŀѡ��)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "����,1200,0,;����,2700,0,", Me.Name & "\������Ŀѡ��", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                Else
                    '��������
                    gstrSQL = GetPublicSQL(SQL.��������ѡ��)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "S")
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\��������ѡ��", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                End If

                If bytRet = 1 Then
                    If Index = 0 Then
                        If mclsVsfBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                    Else
                        If mclsVsfAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                    End If
                    
                    .EditText = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("����").Value)
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    DataChanged = True
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 1                                '��������
            
            If Col = .ColIndex("��������") Then

                If Val(Left(.TextMatrix(Row, .ColIndex("���뷽ʽ")), 1)) = 1 Then
                    '���Ʊ���
                    gstrSQL = GetPublicSQL(SQL.������Ŀѡ��)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "����,1200,0,;����,2700,0,", Me.Name & "\������Ŀѡ��", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                Else
                    '��������
                    gstrSQL = GetPublicSQL(SQL.��������ѡ��)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "S")
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\��������ѡ��", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                End If

                If bytRet = 1 Then
                    If Index = 0 Then
                        If mclsVsfBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                    Else
                        If mclsVsfAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                    End If
                    
                    .EditText = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("����").Value)
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    DataChanged = True
                End If
            End If
        End Select
    End With
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        Call mclsVsfBefore.KeyDown(KeyCode, Shift)
    Case 1
        Call mclsVsfAfter.KeyDown(KeyCode, Shift)
    End Select
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
                    
            Select Case Index
            Case 0
                If Col = .ColIndex("��������") Then
                    
                    If Val(Left(.TextMatrix(Row, .ColIndex("���뷽ʽ")), 1)) = 1 Then
                        gstrSQL = GetPublicSQL(SQL.������Ŀ����, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,2700,0,", Me.Name & "\������Ŀ����", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    Else
                        gstrSQL = GetPublicSQL(SQL.�����������, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp, "S")
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\�����������", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    End If

                    If bytRet = 1 Then
                        
                        If Index = 0 Then
                            If mclsVsfBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                Exit Sub
                            End If
                        Else
                            If mclsVsfAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                Exit Sub
                            End If
                        End If
    
                        .EditText = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("����").Value)
                        
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                        DataChanged = True
    
                    Else
                        KeyCode = 0

                        .Cell(flexcpData, Row, Col) = strText
                        .TextMatrix(Row, Col) = strText
    
                    End If
                End If
            Case 1
                If Col = .ColIndex("��������") Then
                    
                    If Val(Left(.TextMatrix(Row, .ColIndex("���뷽ʽ")), 1)) = 1 Then
                        gstrSQL = GetPublicSQL(SQL.������Ŀ����, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,2700,0,", Me.Name & "\������Ŀ����", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    Else
                        gstrSQL = GetPublicSQL(SQL.�����������, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp, "S")
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\�����������", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    End If

                    If bytRet = 1 Then
                        
                        If Index = 0 Then
                            If mclsVsfBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                Exit Sub
                            End If
                        Else
                            If mclsVsfAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                Exit Sub
                            End If
                        End If
    
                        .EditText = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("����").Value)
                        
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                        DataChanged = True
    
                    Else
                        KeyCode = 0

                        .Cell(flexcpData, Row, Col) = strText
                        .TextMatrix(Row, Col) = strText
    
                    End If
                End If
            End Select
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsfBefore.KeyPress(KeyAscii)
    Case 1
        Call mclsVsfAfter.KeyPress(KeyAscii)
    End Select
    
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsfBefore.KeyPressEdit(KeyAscii)
    Case 1
        Call mclsVsfAfter.KeyPressEdit(KeyAscii)
    End Select
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Select Case Index
        Case 0
            Call mclsVsfBefore.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        Case 1
            Call mclsVsfAfter.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        End Select
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsfBefore.EditSelAll
    Case 1
        Call mclsVsfAfter.EditSelAll
    End Select
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsfBefore.BeforeEdit(Row, Col, Cancel)
    Case 1
        Call mclsVsfAfter.BeforeEdit(Row, Col, Cancel)
    End Select
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Select Case Index
    Case 0
        Call mclsVsfBefore.ValidateEdit(Col, Cancel)
    Case 1
        Call mclsVsfAfter.ValidateEdit(Col, Cancel)
    End Select
End Sub

