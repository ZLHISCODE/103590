VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmFeeGroupRollingCurtain 
   BorderStyle     =   0  'None
   Caption         =   "�������տ����"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picCurrentMoney 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      ScaleHeight     =   225
      ScaleWidth      =   2265
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
      Begin VB.Label lblCurrentMoney 
         Appearance      =   0  'Flat
         Caption         =   "��ǰ�ݴ��: "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   8355
      End
   End
   Begin VB.PictureBox picSendFeeDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   2280
      ScaleHeight     =   2295
      ScaleWidth      =   6375
      TabIndex        =   9
      Top             =   5160
      Width           =   6375
      Begin VB.PictureBox picImgPlanSub 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   13
         Top             =   30
         Width           =   210
         Begin VB.Image imgColPlanSub 
            Height          =   195
            Left            =   0
            Picture         =   "frmFeeGroupRollingCurtain.frx":0000
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsSubCollectorInfo 
         Height          =   1695
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   5535
         _cx             =   9763
         _cy             =   2990
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeGroupRollingCurtain.frx":054E
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
         ExplorerBar     =   5
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
   Begin VB.PictureBox picTabSendFee 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4920
      ScaleHeight     =   1575
      ScaleWidth      =   5175
      TabIndex        =   7
      Top             =   480
      Width           =   5175
      Begin XtremeSuiteControls.TabControl tabSubSendFee 
         Height          =   1935
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2295
         _Version        =   589884
         _ExtentX        =   4048
         _ExtentY        =   3413
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picLastTime 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   1320
      ScaleHeight     =   2775
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   2640
      Width           =   6975
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   14
         Top             =   510
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmFeeGroupRollingCurtain.frx":076F
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.CommandButton cmdSendFees 
         Caption         =   "����(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   4
         Top             =   75
         Width           =   1300
      End
      Begin VB.CommandButton cmdReloadData 
         Caption         =   "������ȡ��������(&G)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   75
         Width           =   2355
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCollectHistory 
         Height          =   2055
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   3255
         _cx             =   5741
         _cy             =   3625
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeGroupRollingCurtain.frx":0CBD
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
         ExplorerBar     =   5
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
      Begin MSComCtl2.DTPicker dtpLastTime 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   75
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116719617
         CurrentDate     =   41521
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   300
         Left            =   5160
         TabIndex        =   2
         Top             =   75
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116719617
         CurrentDate     =   41521
      End
      Begin VB.Label lblLastTime 
         AutoSize        =   -1  'True
         Caption         =   "�ϴ�����ʱ��                           ��ֹʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4935
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1320
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpSendFees 
      Bindings        =   "frmFeeGroupRollingCurtain.frx":0E4D
      Left            =   120
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFeeGroupRollingCurtain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjChargeBillRC As New clsChargeBill, mfrmChargeBillTotalRC As Form    '�տ���Ϣ��Ʊ�ݶ���
Private mlngModule As Long, mstrPrivs As String
Private mlngGroupID As Long '�ɿ���ID
Private mcbrPopupSub As CommandBar

Private Enum EM_Tab
    EM_Tab_�տ� = 1
    EM_Tab_���� = 2
    EM_Tab_��ʷ������Ϣ = 3
    EM_Tab_�տƱ�ݻ��� = 4
    EM_Tab_�շ�Ա������ϸ = 5
    EM_Tab_���տ���Ϣ = 6
    EM_Tab_�շ�Ա������Ϣ = 7
End Enum

Private Sub dkpSendFees_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionAttaching Then Cancel = True
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dtpEndTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtpLastTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:������
    '����:2013-09-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
        
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    '��ʼ������
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    
    Set mcbrPopupSub = cbsThis.Add("�����˵�2", xtpBarPopup)
    With mcbrPopupSub.Controls
        .Add xtpControlButton, conMenu_View_Detail, "��ʾ��ϸ"
    End With
    
    cbsThis.ActiveMenuBar.Visible = False
    
    zlDefCommandBars = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ViewDetail()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�鿴��ϸ��ť����
    '����:������
    '����:2013-09-22
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim i As Integer, strIDs As String
    If ActiveControl = vsSubCollectorInfo Then
        With vsSubCollectorInfo
            For i = .Row To .RowSel
                strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
            Next i
            strIDs = Mid(strIDs, 2)
            Call mobjChargeBillRC.ChargeRollingListShow(Me, EM_�շ�Ա����, strIDs, mlngModule, mstrPrivs)
        End With
    Else
        With vsCollectHistory
            For i = .Row To .RowSel
                strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
            Next i
            strIDs = Mid(strIDs, 2)
            Call mobjChargeBillRC.ChargeRollingListShow(Me, EM_С���տ�, strIDs, mlngModule, mstrPrivs)
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub InitMe(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngGroupID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ʽ���
    '���:lngModule-ģ���
    '     strPrivs-Ȩ�޴�
    '����:������
    '����:2013-10-10
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngGroupID = lngGroupID
End Sub

Private Sub SetDockingPanel()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:����DOCKINGPANEL�ؼ�
    '����:������
    '����:2013-09-04
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    
    With dkpSendFees
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(1, 1000, 1000, DockTopOf)
        objPanel.Handle = picLastTime.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.MinTrackSize.Height = 150
        Set objPanel = .CreatePane(2, 1000, 1000, DockBottomOf, objPanel)
        objPanel.Handle = picTabSendFee.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.MinTrackSize.Height = 250
        Set objPanel = .CreatePane(3, 1000, 300, DockBottomOf, objPanel)
        objPanel.Handle = picCurrentMoney.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.MinTrackSize.Height = 35
        objPanel.MaxTrackSize.Height = 35
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_Detail
            Call ViewDetail
    End Select
End Sub

Private Sub cmdReloadData_Click()
    Call SetDefaultRollingCurtain(True)
End Sub

Private Sub cmdSendFees_Click()
    Call RollingCurtain
End Sub

Private Sub SaveRollingCurtain()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�������ʲ���
    '����:������
    '����:2013-09-10
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNO As String, strSQL As String, strIDs As String, i As Integer, lngID As Long
    Dim strTemp As String, blnBatch As Boolean, colSql As New Collection, strFixedSql As String
    blnBatch = False
    
    '��ȡ���ݺ���ID
    strNO = zlDatabase.GetNextNo(139)
    lngID = zlDatabase.GetNextId("��Ա�սɼ�¼")
    strFixedSql = "Zl_С�����ʼ�¼_Insert(" & lngID & ",'" & strNO & "'," & mlngGroupID & "," & _
                  "to_date('" & dtpLastTime.Value & "','yyyy-MM-dd HH24:mi:ss'),to_date('" & dtpEndTime.Value & "','yyyy-MM-dd HH24:mi:ss'),'" & _
                   UserInfo.���� & "'," & UserInfo.ID & ",to_date('" & zlDatabase.Currentdate & "','yyyy-MM-dd HH24:mi:ss'),'"
    
    With vsCollectHistory
        For i = 1 To .Rows - 1
            strTemp = strIDs
            strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
            If Len(strIDs) >= 4000 Then
                strTemp = Mid(strTemp, 2)
                If blnBatch = False Then
                    strSQL = strFixedSql & strTemp & "',1)"
                Else
                    strSQL = strFixedSql & strTemp & "',2)"
                End If
                blnBatch = True
                colSql.Add strSQL
                strIDs = "," & Val(.TextMatrix(i, .ColIndex("ID")))
            End If
        Next i
    End With
    
    strIDs = Mid(strIDs, 2)
    If strIDs <> "" Then
        If blnBatch = False Then
            strSQL = strFixedSql & strIDs & "',0)"
        Else
            strSQL = strFixedSql & strIDs & "',3)"
        End If
        colSql.Add strSQL
    End If
    
    On Error GoTo errSql
    Call zlExecuteProcedureArrAy(colSql, Me.Caption)
    Call frmFeeGroupManage.AutoPrint(lngID, strNO, 2)
    Exit Sub
errSql:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub SetDefaultRollingCurtain(ByVal blnReload As Boolean, Optional ByVal blnUpdateEndTime As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:����Ĭ�����ʽ�����Ϣ
    '����:������
    '����:2013-09-10
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strSQL As String, rsTmp As New ADODB.Recordset, i As Integer, strDate As String
    Dim strEndDate As String
    If blnReload = True Then GoTo Reload
    strDate = ""
    
    If strDate = "" Then
        strSQL = "Select �ϴ�����ʱ�� From �������鳤���� Where ��Id= [1] And �鳤Id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, UserInfo.ID)
        If rsTmp.RecordCount <> 0 Then
            strDate = Nvl(rsTmp!�ϴ�����ʱ��)
        End If
    End If
    If strDate = "" Then
        strSQL = "Select �ϴ�����ʱ�� From ����ɿ���� Where Id= [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
        If rsTmp.RecordCount <> 0 Then
            strDate = Nvl(rsTmp!�ϴ�����ʱ��)
        End If
    End If
    If strDate = "" Then
        strSQL = "Select ��ֹʱ�� From ��Ա�սɼ�¼ Where ��¼����=3 And ����ʱ�� Is Null And �ɿ���ID= [1] Order By ��ֹʱ�� desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
        If rsTmp.RecordCount <> 0 Then
            strDate = Nvl(rsTmp!��ֹʱ��)
        End If
    End If
    If strDate = "" Then
        strSQL = "Select �Ǽ�ʱ�� From ��Ա�սɼ�¼ Where ��¼����=2 And ����ʱ�� Is Null And �ɿ���ID= [1] Order By �Ǽ�ʱ�� asc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
        If rsTmp.RecordCount <> 0 Then
            strDate = Nvl(rsTmp!�Ǽ�ʱ��)
        End If
    End If
    
    If strDate = "" Then
        dtpLastTime.Enabled = True
        strDate = Format(DateAdd("d", -7, zlDatabase.Currentdate), "yyyy-mm-dd HH:MM:SS")
    End If
    dtpLastTime.Value = strDate
Reload:
    With vsCollectHistory
        .Rows = 1
        strEndDate = zlDatabase.Currentdate
        dtpEndTime.MaxDate = strEndDate
        If blnUpdateEndTime Then dtpEndTime.Value = strEndDate
        If CStr(dtpLastTime.Value) <> "" Then
            strSQL = "" & _
            "Select NO, �տ�Ա, �Ǽ�ʱ��, Trim(to_char(��Ԥ����,'99999999990.00')) As ��Ԥ����, Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�, " & _
            "Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�, С���տ���, С���տ�ʱ��, ժҪ, ID" & vbNewLine & _
            "From ��Ա�սɼ�¼" & vbNewLine & _
            "Where ��¼���� = 2 And С���տ��� = [1] And ����ʱ�� Is Null" & vbNewLine & _
            "      And С���տ�ʱ�� Between [2] And [3] And С������ID Is Null And �ɿ���ID = [4] " & vbNewLine & _
            "Order By �Ǽ�ʱ��,NO Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, CDate(dtpLastTime.Value), CDate(dtpEndTime.Value), mlngGroupID)
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("�տ��")) = Nvl(rsTmp!NO)
                .TextMatrix(.Rows - 1, .ColIndex("�տ�ʱ��")) = Nvl(rsTmp!�Ǽ�ʱ��)
                .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = Nvl(rsTmp!��Ԥ����)
                .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = Nvl(rsTmp!����ϼ�)
                .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = Nvl(rsTmp!����ϼ�)
                .TextMatrix(.Rows - 1, .ColIndex("�ɿ���")) = Nvl(rsTmp!�տ�Ա)
                .TextMatrix(.Rows - 1, .ColIndex("С���տ���")) = Nvl(rsTmp!С���տ���)
                .TextMatrix(.Rows - 1, .ColIndex("С���տ�ʱ��")) = Nvl(rsTmp!С���տ�ʱ��)
                .TextMatrix(.Rows - 1, .ColIndex("��ע")) = Nvl(rsTmp!ժҪ)
                .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTmp!ID)
                rsTmp.MoveNext
            Loop
            'Set .DataSource = rsTmp
            .AutoSize 1, .Cols - 1
            zl_vsGrid_Para_Restore mlngModule, vsCollectHistory, Me.Caption, "С���տ���Ϣ", False
        End If
        If .Rows = 1 Then .Rows = 2
    End With
    With vsSubCollectorInfo
        .Clear 1
        .Rows = 2
    End With
    With vsCollectHistory
        mobjChargeBillRC.LoadChargeAndBillTotalData Me, mlngModule, mstrPrivs, EM_С���տ�, 0
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CancelCollect()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�տ����ϲ���
    '����:������
    '����:2013-09-10
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNO As String, strSQL As String
    With vsCollectHistory
        strSQL = "Zl_С���տ��¼_Cancel(" & Val(.TextMatrix(.RowSel, .ColIndex("ID"))) & ",'" & UserInfo.���� & _
                 "',to_date('" & zlDatabase.Currentdate & "','yyyy-MM-dd HH24:mi:ss'))"
    End With
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshCurrentMoney(ByVal intPanel As Integer)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:ˢ�½����ݴ��
    '���:intPanel-TAB�������
    '����:������
    '����:2013-09-18
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select ���㷽ʽ,��� From ��Ա�ɿ���� Where �տ�Ա=[1] And ����=1"
    If intPanel = 1 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����)
    End If
    lblCurrentMoney(intPanel).Caption = " ��ǰ�ݴ��:   "
    If rsTmp.RecordCount <> 0 Then
        Do While Not rsTmp.EOF
            If Val(Nvl(rsTmp!���)) <> 0 Then
                lblCurrentMoney(intPanel).Caption = lblCurrentMoney(intPanel).Caption & rsTmp!���㷽ʽ & ":" & rsTmp!��� & "Ԫ   "
            End If
            rsTmp.MoveNext
        Loop
    End If
    If intPanel = 1 Then
        vsCollectHistory.Select 0, 0
        vsSubCollectorInfo.Rows = 2
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub ButtonCancelCollect()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�շ����ϰ�ť����
    '����:������
    '����:2013-09-22
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With vsCollectHistory
        If MsgBox("���տ��¼[" & .TextMatrix(.RowSel, .ColIndex("�տ��")) & "]���ϣ�ȷ�����ϣ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End With
    
    Call CancelCollect
    Call SetDefaultRollingCurtain(True)
    Call RefreshCurrentMoney(1)
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub RollingCurtain()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:���ʰ�ť����
    '����:������
    '����:2013-09-22
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNOs As String, i As Integer
    With vsCollectHistory
        For i = 1 To .Rows - 1
            strNOs = strNOs & "," & .TextMatrix(i, .ColIndex("�տ��"))
        Next i
        strNOs = Mid(strNOs, 2)
    End With
    If MsgBox("�Ƿ�������µ��տ�ݽ������ʣ�" & vbCrLf & strNOs, vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Call SaveRollingCurtain
    Call SetDefaultRollingCurtain(False, True)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call zl_vsGrid_Para_Save(mlngModule, vsSubCollectorInfo, Me.Caption, "�շ�Ա������ϸ", False)
    If Not mfrmChargeBillTotalRC Is Nothing Then Unload mfrmChargeBillTotalRC
    Set mobjChargeBillRC = Nothing
End Sub

Private Sub picCurrentMoney_Resize()
    On Error Resume Next
    With lblCurrentMoney(1)
        .Top = 15
        .Width = picCurrentMoney.Width - 15
        .Height = picCurrentMoney.Height - 15
    End With
End Sub

Private Sub vsCollectHistory_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strSQL As String, rsTmp As New ADODB.Recordset, i As Integer
    If OldRow = NewRow Then Exit Sub
    With vsCollectHistory
        'If .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then .Select 0, 0
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
    End With
    
    With vsSubCollectorInfo
        .Rows = 1
        strSQL = "Select No As ���ʵ���, �Ǽ�ʱ�� As ����ʱ��, ��ʼʱ�� As ��ʼʱ��, ��ֹʱ�� As ��ֹʱ��, " & _
                 "       Trim(to_char(��Ԥ����,'99999999990.00')) As ��Ԥ����, Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�," & _
                 "       Trim(to_char(����ϼ�,'99999999990.00')) As ����ϼ�, �Ǽ���, �Ǽ�ʱ��, С���տ���, С���տ�ʱ��, ժҪ As ��ע, ID, �տ�Ա" & vbNewLine & _
                 "From ��Ա�սɼ�¼ " & vbNewLine & _
                 "Where ��¼���� = 1 And ����ʱ�� Is Null And С���տ�ID= [1]" & vbNewLine & _
                 "Order By ����ʱ�� Desc"
        With vsCollectHistory
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.RowSel, .ColIndex("ID"))))
        End With
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("���ʵ���")) = Nvl(rsTmp!���ʵ���)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = Nvl(rsTmp!����ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ʼʱ��")) = Nvl(rsTmp!��ʼʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ֹʱ��")) = Nvl(rsTmp!��ֹʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = Nvl(rsTmp!��Ԥ����)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = Nvl(rsTmp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = Nvl(rsTmp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("�Ǽ���")) = Nvl(rsTmp!�Ǽ���)
            .TextMatrix(.Rows - 1, .ColIndex("�Ǽ�ʱ��")) = Nvl(rsTmp!�Ǽ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ���")) = Nvl(rsTmp!С���տ���)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ�ʱ��")) = Nvl(rsTmp!С���տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ע")) = Nvl(rsTmp!��ע)
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTmp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("�տ�Ա")) = Nvl(rsTmp!�տ�Ա)
            rsTmp.MoveNext
        Loop
        'Set .DataSource = rsTmp
        .AutoSize 1, .Cols - 1
        zl_vsGrid_Para_Restore mlngModule, vsSubCollectorInfo, Me.Caption, "�շ�Ա������ϸ", False
        With vsCollectHistory
            mobjChargeBillRC.LoadChargeAndBillTotalData Me, mlngModule, mstrPrivs, EM_С���տ�, Val(.TextMatrix(.RowSel, .ColIndex("ID")))
        End With
        If .Rows = 1 Then .Rows = 2
    End With
    Call zl_VsGridRowChange(vsCollectHistory, OldRow, NewRow, OldCol, NewCol)
    With vsCollectHistory
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Public Sub ClearChargeAndBillTotalForm()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�ⲿ�������Ʊ�ݴ�������
    '����:������
    '����:2013-10-12
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    mobjChargeBillRC.ClearChargeAndBillTotalForm
End Sub

Private Sub SetTabControl()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:����TAB�ؼ�
    '����:������
    '����:2013-09-04
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With tabSubSendFee
        Set .PaintManager.Font = lblLastTime.Font
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        .InsertItem EM_Tab_�տƱ�ݻ���, " �տƱ�ݻ���  ", mfrmChargeBillTotalRC.hWnd, 0
        .InsertItem EM_Tab_�շ�Ա������ϸ, " �շ�Ա������ϸ  ", picSendFeeDetail.hWnd, 0
        .Item(0).Selected = True
        .PaintManager.BoldSelected = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetDateUnit()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�������ڿؼ���ʽ����
    '����:������
    '����:2013-09-09
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    dtpLastTime.Format = dtpCustom
    dtpLastTime.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpEndTime.Format = dtpCustom
    dtpEndTime.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpEndTime.Value = dtpEndTime.Value + 1
End Sub

Private Sub picTabSendFee_Resize()
    On Error Resume Next
    tabSubSendFee.Width = picTabSendFee.Width
    tabSubSendFee.Height = picTabSendFee.Height
End Sub

Private Sub picLastTime_Resize()
    On Error Resume Next
    cmdSendFees.Left = picLastTime.Width - cmdSendFees.Width - 300
    cmdReloadData.Left = cmdSendFees.Left - cmdReloadData.Width - 300
    If cmdReloadData.Left < dtpEndTime.Left + dtpEndTime.Width + 200 Then
        cmdReloadData.Left = dtpEndTime.Left + dtpEndTime.Width + 200
        cmdSendFees.Left = cmdReloadData.Left + cmdReloadData.Width + 300
    End If
    With vsCollectHistory
        .Width = picLastTime.Width - 15
        .Height = picLastTime.Height - 430
    End With
End Sub

Private Sub vsCollectHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsCollectHistory, Me.Caption, "С���տ���Ϣ", False)
End Sub

Private Sub vsCollectHistory_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsCollectHistory_DblClick()
    With vsCollectHistory
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        Call mobjChargeBillRC.ChargeRollingListShow(Me, EM_С���տ�, Val(.TextMatrix(.RowSel, .ColIndex("ID"))), mlngModule, mstrPrivs)
    End With
End Sub

Private Sub vsCollectHistory_GotFocus()
    Call zl_VsGridGotFocus(vsCollectHistory)
    With vsCollectHistory
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsCollectHistory_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsCollectHistory)
End Sub

Private Sub vsSubCollectorInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call zl_VsGridRowChange(vsSubCollectorInfo, OldRow, NewRow, OldCol, NewCol)
    With vsSubCollectorInfo
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsSubCollectorInfo_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsSubCollectorInfo, Me.Caption, "�շ�Ա������ϸ", False)
End Sub

Private Sub vsSubCollectorInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsSubCollectorInfo_DblClick()
    With vsSubCollectorInfo
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        Call mobjChargeBillRC.ChargeRollingListShow(Me, EM_�շ�Ա����, Val(.TextMatrix(.RowSel, .ColIndex("ID"))), mlngModule, mstrPrivs)
    End With
End Sub

Private Sub vsSubCollectorInfo_GotFocus()
    Call zl_VsGridGotFocus(vsSubCollectorInfo)
    With vsSubCollectorInfo
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsSubCollectorInfo_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsSubCollectorInfo)
End Sub

Private Sub vsSubCollectorInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intRow As Integer
    With vsSubCollectorInfo
        If .TextMatrix(1, .ColIndex("ID")) = "" Then Exit Sub
        If Button = 2 Then
            If y <= 255 Then
                Exit Sub
            End If
            intRow = y \ 255
            If intRow > .Rows - 1 Then Exit Sub
            If .Enabled And .Visible Then .SetFocus
            .Select intRow, 0
            mcbrPopupSub.ShowPopup
        End If
    End With
End Sub

Private Sub SetGrid()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSF�ؼ�
    '����:������
    '����:2013-10-13
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    With vsSubCollectorInfo
        For i = 0 To .Cols - 1
            If .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "�տ�Ա" Or .ColKey(i) = "����" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "���ʵ���" Or .ColKey(i) = "��ʼʱ��" Or .ColKey(i) = "��ֹʱ��" Then .ColData(i) = "1|0"
        Next
    End With
    
    With vsCollectHistory
        For i = 0 To .Cols - 1
            If .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "С���տ���" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "����" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "�տ��" Or .ColKey(i) = "�տ�ʱ��" Then .ColData(i) = "1|0"
        Next
    End With
    zl_vsGrid_Para_Restore mlngModule, vsSubCollectorInfo, Me.Caption, "�շ�Ա������ϸ", False
    zl_vsGrid_Para_Restore mlngModule, vsCollectHistory, Me.Caption, "С���տ���Ϣ", False
End Sub

Public Sub RefreshPage()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:ˢ�½���
    '����:������
    '����:2013-10-13
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Call SetDefaultRollingCurtain(True, True)
    Call RefreshCurrentMoney(1)
    vsCollectHistory.Select 0, 0
End Sub

Private Sub picSendFeeDetail_Resize()
    On Error Resume Next
    With vsSubCollectorInfo
        .Width = picSendFeeDetail.Width
        .Height = picSendFeeDetail.Height
    End With
End Sub

Private Sub dtpEndTime_Change()
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    strSQL = "Select �ϴ�����ʱ�� From �������鳤���� Where ��Id= [1] And �鳤Id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, UserInfo.ID)
    If rsTmp.RecordCount = 0 Then
        If IsNull(rsTmp!�ϴ�����ʱ��) Then
            strSQL = "Select �ϴ�����ʱ�� From ����ɿ���� Where Id= [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
        End If
    End If
    '�޼�¼�����ϴ�����ʱ�����˳�
    If rsTmp.RecordCount = 0 Then Exit Sub
    If IsNull(rsTmp!�ϴ�����ʱ��) Then Exit Sub
    If dtpEndTime.Value <= CDate(rsTmp!�ϴ�����ʱ��) Then
        dtpEndTime.Value = rsTmp!�ϴ�����ʱ��
    End If
End Sub

Private Sub Form_Load()
    mobjChargeBillRC.SetFontSize lblCurrentMoney(1).Font.Size
    Set mfrmChargeBillTotalRC = mobjChargeBillRC.GetChargeAndBillTotalForm
    cmdSendFees.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
    Call zlDefCommandBars
    Call SetDockingPanel
    Call SetTabControl
    Call SetDateUnit
    Call SetGrid
    '���ʽ���Ĭ����Ϣ
    Call SetDefaultRollingCurtain(False)
End Sub

Private Sub imgColPlanSub_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanSub.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlanSub.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsSubCollectorInfo, lngLeft, lngTop, imgColPlanSub.Height)
    zl_vsGrid_Para_Save mlngModule, vsSubCollectorInfo, Me.Caption, "�շ�Ա������ϸ", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlanSub_Click()
    Call imgColPlanSub_Click
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCollectHistory, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsCollectHistory, Me.Caption, "С���տ���Ϣ", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
