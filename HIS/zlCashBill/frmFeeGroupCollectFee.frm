VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmFeeGroupCollectFee 
   BorderStyle     =   0  'None
   Caption         =   "�������տ����"
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picCurrentMoney 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      ScaleHeight     =   345
      ScaleWidth      =   5865
      TabIndex        =   6
      Top             =   480
      Width           =   5895
      Begin VB.Label lblCurrentMoney 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "��ǰ�ݴ��:    �ֽ�:3000Ԫ    ֧Ʊ:5000Ԫ    ҽ������:10000Ԫ    �����˻�:100Ԫ"
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
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   8295
      End
   End
   Begin VB.PictureBox picSubWorker 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   3000
      ScaleHeight     =   4215
      ScaleWidth      =   1935
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
      Begin MSComctlLib.ListView lvwSubWorker_S 
         Height          =   4335
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   7646
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilsWorker"
         SmallIcons      =   "ilsWorkerSmall"
         ColHdrIcons     =   "ilsWorkerSmall"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "C1"
            Text            =   "����"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "C2"
            Text            =   "���"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "C3"
            Text            =   "����"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Key             =   "C4"
            Text            =   "��������"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.PictureBox picGeneralInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   5760
      ScaleHeight     =   2655
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   4440
      Width           =   3735
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   8
         Top             =   450
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmFeeGroupCollectFee.frx":0000
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.TextBox txtSendNO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   2
         Top             =   8
         Width           =   2500
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCollectorInfo 
         Height          =   1095
         Left            =   0
         TabIndex        =   3
         Top             =   420
         Width           =   2655
         _cx             =   4683
         _cy             =   1931
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
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeGroupCollectFee.frx":054E
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
         Editable        =   2
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
      Begin VB.Label lblInfo 
         Caption         =   "���ʵ���"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   60
         Width           =   840
      End
   End
   Begin MSComctlLib.ImageList ilsWorker 
      Left            =   480
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFeeGroupCollectFee.frx":073B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFeeGroupCollectFee.frx":1015
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsWorkerSmall 
      Left            =   1200
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFeeGroupCollectFee.frx":18EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFeeGroupCollectFee.frx":1E89
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpCollectFees 
      Left            =   480
      Top             =   600
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFeeGroupCollectFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjChargeBillCollect As New clsChargeBill, mfrmChargeBillTotalCollect As Form  '�տ���Ϣ��Ʊ�ݶ���
Private mlngModule As Long, mstrPrivs As String, mintColumn As Integer
Private mlngGroupID As Long '�ɿ���ID
Private mfrmMain As Form    '������
Private mcbrListView As CommandBar, mcbrControl As CommandBarControl

Private Enum EM_Pan
    EM_Pan_��Ա�� = 1
    EM_Pan_�շ�������Ϣ = 2
    EM_Pan_�տƱ����Ϣ = 3
    EM_Pan_��Ա��� = 4
End Enum

Private Enum EM_SPM
    EM_SPM_�տ��б� = 1
    EM_SPM_��Ա�б� = 2
End Enum

Public Event ShowPopupMenu(ByVal bytType As Byte) 'bytType-1:�������տ��б����˵�;2:��������Ա�б����˵�

Private Sub SetGrid()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSF�ؼ�
    '����:������
    '����:2013-10-13
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    With vsCollectorInfo
        .Cell(flexcpChecked, 0, .ColIndex("ѡ��")) = flexUnchecked
        For i = 0 To .Cols - 1
            If .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "�տ�Ա" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "����" Or .ColKey(i) = "�տ��" Or .ColKey(i) = "ѡ��" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "���ʵ���" Or .ColKey(i) = "��ʼʱ��" Or .ColKey(i) = "��ֹʱ��" Then .ColData(i) = "1|0"
        Next
    End With
    
    zl_vsGrid_Para_Restore mlngModule, vsCollectorInfo, Me.Caption, "�շ�Ա������Ϣ", False

End Sub

Public Sub ClearChargeAndBillTotalForm()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�ⲿ�������Ʊ�ݴ�������
    '����:������
    '����:2013-10-12
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Call mobjChargeBillCollect.ClearChargeAndBillTotalForm
End Sub

Public Sub ChargeRollingListShow(frmMain As Object, bytType As TotalType, strIDs As String)
    Call mobjChargeBillCollect.ChargeRollingListShow(frmMain, bytType, strIDs, mlngModule, mstrPrivs)
End Sub

Public Sub InitMe(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngGroupID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���տ����
    '���:frmMain-������
    '     lngModule-ģ���
    '     strPrivs-Ȩ�޴�
    '     lngGroupID-��ID
    '����:������
    '����:2013-10-10
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    mlngModule = lngModule
    mlngGroupID = lngGroupID
    mstrPrivs = strPrivs
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmChargeBillTotalCollect Is Nothing Then Unload mfrmChargeBillTotalCollect
    Set mobjChargeBillCollect = Nothing
    SaveWinState Me, App.ProductName, "frmFeeGroupCollectFee"
End Sub

Private Sub lvwSubWorker_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        RaiseEvent ShowPopupMenu(EM_SPM_��Ա�б�)
    End If
End Sub

Private Sub picCurrentMoney_Resize()
    On Error Resume Next
    With lblCurrentMoney(0)
        .Top = 15
        .Width = picCurrentMoney.Width - 15
        .Height = picCurrentMoney.Height - 15
    End With
End Sub

Private Sub txtSendNO_GotFocus()
    Call zlControl.TxtSelAll(txtSendNO)
End Sub

Private Sub dkpCollectFees_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionAttaching Then Cancel = True
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub txtSendNO_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandle
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 65 Or KeyAscii > 90) And _
       (KeyAscii < 97 Or KeyAscii > 122) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        If txtSendNO.Text = "" Then
            KeyAscii = 0
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        Dim i As Integer, strSQL As String
        Dim rsTmp As New ADODB.Recordset
        '��ȫƥ�����뵥��
        With vsCollectorInfo
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ʵ���")) = txtSendNO.Text Then
                    If .Enabled And .Visible Then .SetFocus
                    DoEvents
                    .Select i, .ColIndex("ѡ��")
                    .TopRow = i
                    Exit Sub
                End If
            Next i
            strSQL = "Select �տ�Ա" & vbNewLine & _
                     "From ��Ա�սɼ�¼" & vbNewLine & _
                     "Where ��¼���� = 1 And �ɿ���id = [1] And (С���տ��� = [3] Or С���տ��� Is Null) And ����ʱ�� Is Null And С���տ�id Is Null And NO = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, txtSendNO.Text, UserInfo.����)
            If rsTmp.RecordCount <> 0 Then
                LoadWorkerCollectDetail (NVL(rsTmp!�տ�Ա))
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("���ʵ���")) = txtSendNO.Text Then
                        If .Enabled And .Visible Then .SetFocus
                        DoEvents
                        .Select i, 1
                        .TopRow = i
                        Exit Sub
                    End If
                Next i
            End If
        End With
        
        '�Զ��������뵥��,�ٴν��в���
        txtSendNO.Text = GetFullNO(txtSendNO.Text, 137)
        With vsCollectorInfo
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ʵ���")) = txtSendNO.Text Then
                    If .Enabled And .Visible Then .SetFocus
                    DoEvents
                    .Select i, .ColIndex("ѡ��")
                    .TopRow = i
                    Exit Sub
                End If
            Next i
            strSQL = "Select �տ�Ա" & vbNewLine & _
                     "From ��Ա�սɼ�¼" & vbNewLine & _
                     "Where ��¼���� = 1 And �ɿ���id = [1] And ����ʱ�� Is Null And (С���տ��� = [3] Or С���տ��� Is Null) And С���տ�id Is Null And NO = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, txtSendNO.Text, UserInfo.����)
            If rsTmp.RecordCount <> 0 Then
                LoadWorkerCollectDetail (NVL(rsTmp!�տ�Ա))
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("���ʵ���")) = txtSendNO.Text Then
                        If .Enabled And .Visible Then .SetFocus
                        DoEvents
                        .Select i, 1
                        .TopRow = i
                        Exit Sub
                    End If
                Next i
            End If
        End With
        MsgBox "û���ҵ����ʵ���[" & txtSendNO.Text & "]�ļ�¼��", vbInformation, gstrSysName
        If txtSendNO.Visible Then txtSendNO.SetFocus
        Call zlControl.TxtSelAll(txtSendNO)
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
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
    With dkpCollectFees
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(EM_Pan_��Ա��, 300, 1800, DockLeftOf)
        objPanel.Handle = picSubWorker.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
        objPanel.MinTrackSize.Width = 75
        objPanel.MaxTrackSize.Width = 300
        Set objPanel = .CreatePane(EM_Pan_�շ�������Ϣ, 2000, 800, DockRightOf, objPanel)
        objPanel.Handle = picGeneralInfo.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
        objPanel.Title = "�շ�������Ϣ"
        objPanel.MinTrackSize.Height = 100
        Set objPanel = .CreatePane(EM_Pan_�տƱ����Ϣ, 2000, 1000, DockBottomOf, objPanel)
        objPanel.Handle = mfrmChargeBillTotalCollect.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.Title = "�տƱ����Ϣ"
        objPanel.MinTrackSize.Height = 230
        Set objPanel = .CreatePane(EM_Pan_��Ա���, 2000, 100, DockBottomOf)
        objPanel.Handle = picCurrentMoney.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.Title = "��Ա���"
        objPanel.MinTrackSize.Height = 35
        objPanel.MaxTrackSize.Height = 35
        Set .PaintManager.CaptionFont = lblCurrentMoney(0).Font
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picSubWorker_Resize()
    On Error Resume Next
    lvwSubWorker_S.Width = picSubWorker.Width
    lvwSubWorker_S.Height = picSubWorker.Height
End Sub

Public Sub ChangeListViewType(ByVal intTYPE As Integer)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:������Ա�б���ʾ��ʽ
    '���:intType-�б���ʾ��ʽ: 1-��ͼ��;2-Сͼ��;3-�б�;4-��ϸ�б�
    '����:������
    '����:2013-10-09
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Select Case intTYPE
        Case 1
            lvwSubWorker_S.View = lvwIcon
        Case 2
            lvwSubWorker_S.View = lvwSmallIcon
        Case 3
            lvwSubWorker_S.View = lvwList
        Case 4
            lvwSubWorker_S.View = lvwReport
    End Select
End Sub

Private Sub picGeneralInfo_Resize()
    On Error Resume Next
    With vsCollectorInfo
        .Width = picGeneralInfo.Width - 15
        .Height = picGeneralInfo.Height - 430
    End With
End Sub

Private Sub lvwSubWorker_S_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    lvwSubWorker_S.Drag 0
End Sub

Private Sub lvwSubWorker_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LoadWorkerCollectDetail(Item.Text)
End Sub

Private Sub lvwSubWorker_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '��ֹ�༭�����ƶ���Ա�б�
    If Button = 1 Then
        If lvwSubWorker_S.HitTest(x, y) Is Nothing Then Exit Sub
        lvwSubWorker_S.Drag 1
    End If
End Sub

Public Sub AfterCollectEdit()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:С���տ���Ϻ�ˢ�½�������
    '����:������
    '����:2013-09-12
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Call LoadWorkerCollectDetail(lvwSubWorker_S.SelectedItem.Text)
End Sub

Private Sub LoadWorkerCollectDetail(ByVal strWorker As String)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����շ�Ա���շ���Ϣ
    '���:strWorker--�շ�Ա
    '����:������
    '����:2013-09-09
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strSQL As String, rsTmp As New ADODB.Recordset, i As Integer
    strSQL = "" & _
    "Select a.ID, a.NO, a.�Ǽ�ʱ��,Substr(Decode(�Ƿ�Һ�,1,',�Һ�','') || Decode(�Ƿ���￨,1,',���￨','') || Decode(�Ƿ����ѿ�,1,',���ѿ�','') || Decode(�Ƿ��շ�,1,',�շ�','') || Decode(�Ƿ����,1,',����','') || Decode(Ԥ�����,1,',Ԥ��',2,',����Ԥ��',3,',סԺԤ��',''),2) As �������, " & _
    "       a.��ʼʱ��, a.��ֹʱ��, a.��Ԥ����, a.����ϼ�, a.����ϼ�, a.ժҪ, a.�տ�Ա" & vbNewLine & _
    "From ��Ա�սɼ�¼ A" & vbNewLine & _
    "Where a.��¼���� = 1 And a.�ɿ���id = [1] And (a.С���տ��� = [3] Or a.С���տ��� Is Null) And a.����ʱ�� Is Null And a.С���տ�id Is Null And a.�����տ�ʱ�� Is Null And a.�տ�Ա = [2]" & vbNewLine & _
    "Order by �Ǽ�ʱ�� Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, strWorker, UserInfo.����)
    
    With vsCollectorInfo
        .Rows = 1
        If rsTmp.RecordCount <> 0 Then
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = 0
                '0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�
                .TextMatrix(.Rows - 1, .ColIndex("�������")) = NVL(rsTmp!�������)
                .TextMatrix(.Rows - 1, .ColIndex("���ʵ���")) = NVL(rsTmp!NO)
                .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = NVL(rsTmp!�Ǽ�ʱ��)
                .TextMatrix(.Rows - 1, .ColIndex("�տ�Ա")) = NVL(rsTmp!�տ�Ա)
                '.TextMatrix(.Rows - 1, .ColIndex("�տ��")) = Nvl(rsTmp!��������)
                .TextMatrix(.Rows - 1, .ColIndex("��ʼʱ��")) = NVL(rsTmp!��ʼʱ��)
                .TextMatrix(.Rows - 1, .ColIndex("��ֹʱ��")) = NVL(rsTmp!��ֹʱ��)
                .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = Format(NVL(rsTmp!��Ԥ����), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = Format(NVL(rsTmp!����ϼ�), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = Format(NVL(rsTmp!����ϼ�), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("��ע")) = NVL(rsTmp!ժҪ)
                .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTmp!ID)
                rsTmp.MoveNext
            Loop
            .AutoSize 1, .Cols - 1
            zl_vsGrid_Para_Restore mlngModule, vsCollectorInfo, Me.Caption, "�շ�Ա������Ϣ", False
            .ColWidth(.ColIndex("ѡ��")) = 290
            .ColHidden(.ColIndex("ѡ��")) = False
        End If
        If .Rows = 1 Then .Rows = 2
    End With
    
    Call RefreshCurrentMoney(0)
    mobjChargeBillCollect.ClearChargeAndBillTotalForm
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsCollectorInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    Dim bytMode As Byte
    
    With vsCollectorInfo
        If Col <> .ColIndex("ѡ��") Then Exit Sub
        If .TextMatrix(1, .ColIndex("ID")) <> "" Then
            If Row = 0 Then
                If .Cell(flexcpChecked, 0, .ColIndex("ѡ��")) = flexChecked Or .Cell(flexcpChecked, 0, .ColIndex("ѡ��")) = flexTSChecked Then
                    .Cell(flexcpChecked, 1, .ColIndex("ѡ��"), .Rows - 1) = flexChecked
                Else
                    .Cell(flexcpChecked, 1, .ColIndex("ѡ��"), .Rows - 1) = flexUnchecked
                    
                    If .Cell(flexcpChecked, 0, .ColIndex("ѡ��")) = flexTSGrayed Then .Cell(flexcpChecked, 0, .ColIndex("ѡ��")) = flexUnchecked
                End If
            Else
            
                Call CheckVSFState(bytMode)
                
                If bytMode = 0 Then
                    .Cell(flexcpChecked, 0, .ColIndex("ѡ��")) = flexChecked
                ElseIf bytMode = 1 Then
                    .Cell(flexcpChecked, 0, .ColIndex("ѡ��")) = flexUnchecked
                Else
                    .Cell(flexcpChecked, 0, .ColIndex("ѡ��")) = flexTSGrayed
                End If

            End If
        End If
    End With
End Sub

Private Sub CheckVSFState(ByRef bytMode As Byte)
    '����:�������������б��е�checkbox�Ƿ�ȫ��ѡ�л�ȫ��δѡ��
    '����:bytMode-0:ȫ��ѡ��;1:ȫ��δѡ��;2-����ѡ��
    Dim i As Integer
    Dim blnAllChecked As Boolean, blnAllUnChecked As Boolean
    
    On Error GoTo errHandle
    blnAllChecked = True: blnAllUnChecked = True
    
    With vsCollectorInfo
        
        For i = 1 To .Rows - 1
            Select Case .Cell(flexcpChecked, i, .ColIndex("ѡ��"))
                Case flexChecked
                    blnAllUnChecked = False
                Case flexUnchecked
                    blnAllChecked = False
            End Select
        Next
        
        If blnAllChecked Then
            bytMode = 0
        ElseIf blnAllUnChecked Then
            bytMode = 1
        Else
            bytMode = 2
        End If
    End With
    Exit Sub
errHandle:
If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub vsCollectorInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    With vsCollectorInfo
        'If .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then .Select 0, 0
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        mobjChargeBillCollect.LoadChargeAndBillTotalData Me, mlngModule, mstrPrivs, EM_�շ�Ա����, .TextMatrix(.RowSel, .ColIndex("ID"))
        Call zl_VsGridRowChange(vsCollectorInfo, OldRow, NewRow, OldCol, NewCol)
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsCollectorInfo_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsCollectorInfo, Me.Caption, "�շ�Ա������Ϣ", False)
End Sub

Private Sub vsCollectorInfo_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow < 1 Then Cancel = True
End Sub

Private Sub vsCollectorInfo_DblClick()
    With vsCollectorInfo
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        Call ChargeRollingListShow(mfrmMain, EM_�շ�Ա����, Val(.TextMatrix(.RowSel, .ColIndex("ID"))))
    End With
End Sub

Private Sub vsCollectorInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub vsCollectorInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsCollectorInfo
        If Col <> .ColIndex("ѡ��") Then Cancel = True
        If .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then
            Cancel = True
            Exit Sub
        End If
        .Select Row, .ColIndex("ѡ��")
    End With
End Sub

Private Sub vsCollectorInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsCollectorInfo.ColIndex("ѡ��") Then Cancel = True
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsCollectorInfo_GotFocus()
    With vsCollectorInfo
        If Val(.TextMatrix(1, .ColIndex("ID"))) <> 0 Then
            .Select 1, .ColIndex("ѡ��")
        End If
        Call zl_VsGridGotFocus(vsCollectorInfo)
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
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
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lvwSubWorker_S.SelectedItem.Text)
    End If
    
    lblCurrentMoney(intPanel).Caption = " ��ǰ�ݴ��:   "
    If rsTmp.RecordCount <> 0 Then
        Do While Not rsTmp.EOF
            If Val(NVL(rsTmp!���)) <> 0 Then
                lblCurrentMoney(intPanel).Caption = lblCurrentMoney(intPanel).Caption & rsTmp!���㷽ʽ & ":" & rsTmp!��� & "Ԫ   "
            End If
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function LoadSubWorkers() As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ɿ���������Ա
    '����:mlngGroupID-�ɿ���ID
    '����:�ɹ�����True,ʧ�ܷ���False
    '����:������
    '����:2013-09-03
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim lvwItem As ListItem
    strSQL = "Select ������,������ID From ����ɿ���� Where (ɾ������ Is Null or ɾ������ Between Sysdate And to_date('3000-01-01','YYYY-MM-DD')) And ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
    
    If rsTmp.RecordCount = 0 Then
        LoadSubWorkers = False
        Exit Function
    End If
    
    dkpCollectFees.Panes(1).Title = NVL(rsTmp!������)
    
    strSQL = "Select b.Id, b.���, b.����, b.�Ա�, b.����, d.����" & vbNewLine & _
             "From �ɿ��Ա��� A, ��Ա�� B, ������Ա C, ���ű� D" & vbNewLine & _
             "Where a.��Աid = b.Id And ��id = [1] And a.��Աid = c.��Աid And c.����id = d.Id And c.ȱʡ = 1 " & _
             "Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
    
    Do While Not rsTmp.EOF
        If rsTmp!�Ա� Like "*Ů*" Then
            Set lvwItem = lvwSubWorker_S.ListItems.Add(, "_" & rsTmp!ID, NVL(rsTmp!����), 2, 2)
            lvwItem.SubItems(1) = NVL(rsTmp!���)
            lvwItem.SubItems(2) = NVL(rsTmp!����)
            lvwItem.SubItems(3) = NVL(rsTmp!����)
        Else
            '�л����Ա��������
            Set lvwItem = lvwSubWorker_S.ListItems.Add(, "_" & rsTmp!ID, NVL(rsTmp!����), 1, 1)
            lvwItem.SubItems(1) = NVL(rsTmp!���)
            lvwItem.SubItems(2) = NVL(rsTmp!����)
            lvwItem.SubItems(3) = NVL(rsTmp!����)
        End If
        rsTmp.MoveNext
    Loop
    LoadSubWorkers = True
    Exit Function
errHandle:
    LoadSubWorkers = False
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub lvwSubWorker_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwSubWorker_S.SortOrder = IIf(lvwSubWorker_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwSubWorker_S.SortKey = mintColumn
        lvwSubWorker_S.SortOrder = lvwAscending
    End If
    lvwSubWorker_S.Sorted = True
End Sub

Private Sub Form_Load()
    mobjChargeBillCollect.SetFontSize lblCurrentMoney(0).Font.Size
    Set mfrmChargeBillTotalCollect = mobjChargeBillCollect.GetChargeAndBillTotalForm
    Call SetDockingPanel
    If LoadSubWorkers = False Then
        Call frmFeeGroupManage.FailInit
        Exit Sub
    End If
    Call SetGrid
    vsCollectorInfo.Select 0, 0
    RestoreWinState Me, App.ProductName, "frmFeeGroupCollectFee"
End Sub

Private Sub vsCollectorInfo_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsCollectorInfo)
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCollectorInfo, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsCollectorInfo, Me.Caption, "�շ�Ա������Ϣ", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub

Public Sub SetVSFCheckBat(ByVal bytMode As Byte)
    '����:ȫѡ��ȫ������������б��е�CheckBox
    '����:bytMode-0:ȫѡ;1-ȫ��
    Dim i As Integer
    
    On Error GoTo errHandle
    With vsCollectorInfo
        If .TextMatrix(1, .ColIndex("ID")) = "" Then Exit Sub

        For i = 0 To .Rows - 1
            .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = IIf(bytMode = 0, flexChecked, flexUnchecked)
        Next

    End With
    Exit Sub
errHandle:
If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub vsCollectorInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intRow As Integer
    With vsCollectorInfo
        If .TextMatrix(1, .ColIndex("ID")) = "" Then Exit Sub
        If Button = 2 Then
            If .MouseRow < 1 Then Exit Sub
            If .MouseRow > .Rows - 1 Then Exit Sub
            If .Enabled And .Visible Then .SetFocus
            .Select .MouseRow, 0
            RaiseEvent ShowPopupMenu(EM_SPM_�տ��б�)
        End If
    End With
End Sub


