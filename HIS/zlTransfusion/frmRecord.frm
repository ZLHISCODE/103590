VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecord 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList img16 
      Left            =   2610
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox pictbcKernel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   -15
      ScaleHeight     =   3735
      ScaleWidth      =   6450
      TabIndex        =   0
      Top             =   2505
      Width           =   6450
      Begin XtremeSuiteControls.TabControl tbcKernel 
         Height          =   3495
         Left            =   135
         TabIndex        =   1
         Top             =   150
         Width           =   6210
         _Version        =   589884
         _ExtentX        =   10954
         _ExtentY        =   6165
         _StockProps     =   64
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   1740
      Left            =   60
      TabIndex        =   2
      Top             =   315
      Width           =   6435
      _cx             =   11351
      _cy             =   3069
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   33023
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRecord.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
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
   Begin VB.Label lbl˵�� 
      Alignment       =   1  'Right Justify
      Caption         =   "״̬˵��:  �� ����ִ�� �� ��� �� �ܾ�"
      Height          =   180
      Left            =   90
      TabIndex        =   3
      Top             =   2115
      Width           =   3420
   End
   Begin VB.Image img��� 
      Height          =   240
      Left            =   4905
      Picture         =   "frmRecord.frx":009B
      Top             =   15
      Width           =   240
   End
   Begin VB.Image img��ִ�� 
      Height          =   240
      Left            =   4620
      Picture         =   "frmRecord.frx":68ED
      Top             =   15
      Width           =   240
   End
   Begin VB.Image img�ܾ� 
      Height          =   240
      Left            =   4365
      Picture         =   "frmRecord.frx":D13F
      Top             =   15
      Width           =   240
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   2175
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpSub 
      Bindings        =   "frmRecord.frx":13991
      Left            =   495
      Top             =   30
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum rptCOL
    rptCOL_ִ�з��� = 0
    rptCOL_�ӵ�ʱ�� = 1
    rptCOL_��ҩ�� = 2
    rptCOL_���� = 3
    rptCOL_��ʱ = 4
    rptCOL_��ϵ�� = 5
    rptCOL_�ӵ��� = 6
    rptCOL_��ˮ�� = 7
End Enum


Private Enum vsListCol
    col_������ = 0
    col_��� = 1
    col_˳�� = 2
    col_ҽ������ = 3
    col_���� = 4
    col_��λ = 5
    col_��� = 6
    col_ִ��Ƶ�� = 7
    col_�÷� = 8
    col_Ƥ�Խ�� = 9
    col_�շѽ�� = 10
    col_���� = 11
    col_���� = 12
    col_ʱ�� = 13
    col_���� = 14
    col_ִ���� = 15
    col_�˶��� = 16
    col_ʣ����� = 17
    col_˵�� = 18
    col_BillKey = 19
    col_groupkey = 20
    col_�޸ı�־ = 21
End Enum

Private Const conMenu_File_BillPrintEx As Long = 3554        '��Һƿǩ�����

Private mlngType As Integer '��ʾ���� 0-���� 1-��Һ 2-ע�� 3-Ƥ��

'��Ŀ����
'--------
'Public WithEvents mclsExpenses As zlCISKernel.clsDockExpense
Public WithEvents mclsExpenses As zlPublicExpense.clsDockExpense
Attribute mclsExpenses.VB_VarHelpID = -1
'Private mclsPubExpense As zlPublicExpense.clsPublicExpense
Private mcolSubForm As Collection
Private mfrmActive As Form
'--------
Private mstrGroupKey As String      '��ǰѡ����ִ����Ŀ
Private mlng��ˮ�� As Long          '��ǰѡ������ˮ��
Private mlngModi As Long '�Ƿ��޸�״̬
Private mstrִ���� As String

Private mblnUpdate As Boolean '�Ƿ��޸Ĺ�

Private mfrmMain As frmTransfusion
Private mcbsMain As CommandBars
Private mstrPatiStat As String  '�����嵱ǰ��Ա��״̬
Private marrRecord As Variant   '��ǰִ����Ŀ����

'Private mobjExecRecord As ExecRecord '������Ŀ��

Public Property Get ��ˮ��() As Long
    ��ˮ�� = mlng��ˮ��
End Property

Public Property Let ��ˮ��(ByVal vData As Long)
    mlng��ˮ�� = vData
End Property

Public Property Get ִ����() As String
     ִ���� = mstrִ����
End Property

Public Property Let ִ����(ByVal vData As String)
    mstrִ���� = vData
End Property


Public Property Let �༭(ByVal vData As Long)
    mlngModi = vData
End Property

Public Property Get �༭() As Long
    �༭ = mlngModi
End Property

Public Property Let ��Key(ByVal vData As String)
    mstrGroupKey = vData
End Property

Public Property Get ��Key() As String
    ��Key = mstrGroupKey
End Property

Public Property Let �޸Ĺ�(ByVal vData As Boolean)
    mblnUpdate = vData
End Property

Public Property Get �޸Ĺ�() As Boolean
    �޸Ĺ� = mblnUpdate
End Property

Private Sub cbsSub_Resize()
    On Error Resume Next
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    vsList.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop - lbl˵��.Height
    lbl˵��.Move lngLeft, vsList.Top + vsList.Height, vsList.Width - 45
End Sub

Private Sub dkpSub_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = pictbcKernel.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Call DockSubInit
    pictbcKernel.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
    
    marrRecord = Array()
    
    'TabControl
    '-------------
'    Set mclsExpenses = New zlCISKernel.clsDockExpense
'    Set mclsPubExpense = New zlPublicExpense.clsPublicExpense
    Set mclsExpenses = New zlPublicExpense.clsDockExpense
    Set mcolSubForm = New Collection

    '��ʼ��
'    mclsPubExpense.zlInitCommon glngSys, gcnOracle
    mclsExpenses.zlInitCommon glngSys, gcnOracle

    mcolSubForm.Add mclsExpenses.zlGetForm, "_��Ŀ����"
    With Me.tbcKernel
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        '1257 -��Ŀ���ѹ���
        If GetInsidePrivs(1257, True) <> "" Then
            .InsertItem(0, "��Ŀ����", mcolSubForm("_��Ŀ����").hwnd, 0).Tag = "��Ŀ����"
            .Item(0).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
           ' Call zlDefCommandBars(.Selected) '��ʼˢ�¶���һ�β˵�����ť
        End If
    End With

End Sub

Private Sub DockSubInit()
    Dim objPaneA As Pane, objPaneB As Pane, ojbPaneC As Pane
    Dim lngX As Long
    Dim lngY As Long
    
    'DockingPane ��ʼ��
    '-----------------------------------------------------
    Me.dkpSub.SetCommandBars Me.cbsSub
    
    If GetInsidePrivs(1257, True) <> "" Then
        Set objPaneA = Me.dkpSub.CreatePane(1, 280, 235, DockBottomOf)
        objPaneA.Title = "��Ŀ����"
        objPaneA.Options = PaneNoCloseable Or PaneNoFloatable
    Else
        Me.pictbcKernel.Visible = False
    End If
    
    Me.dkpSub.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpSub.Options.ThemedFloatingFrames = True
    Me.dkpSub.Options.AlphaDockingContext = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    mlng��ˮ�� = 0
    mstrGroupKey = ""
    mlngModi = 0
    mstrִ���� = ""
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    
    Set mclsExpenses = Nothing
    Set marrRecord = Nothing
End Sub

Private Sub mclsExpenses_RequestRefresh()
    Call mfrmMain.ˢ��
End Sub

'Private Sub mclsExpenses_StatusTextUpdate(ByVal Text As String)
'    Call mfrmMain.����״̬��(Text)
'End Sub

Private Sub pictbcKernel_Resize()
    On Error Resume Next
    With pictbcKernel
        tbcKernel.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    '��������ñ������ִ�й���
    
    Dim lng��ˮ�� As Long, lngDeptID As Long, strGroupKey As String
    Dim lngErrNo As Long
    
    Select Case Control.ID
        Case conMenu_Manage_ThingDel
            '�����ӵ�
            lng��ˮ�� = mfrmMain.Get��ˮ��
            If lng��ˮ�� <> 0 Then
                If MsgBox("�Ƿ�����ˮ��Ϊ" & lng��ˮ�� & "��ִ�м�¼��", vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
                    Call mfrmMain.�����ӵ�(lng��ˮ��)
                End If
            End If
        Case conMenu_Edit_Transf_Modify
            '�޸�
            lng��ˮ�� = mfrmMain.Get��ˮ��
            If lng��ˮ�� <> 0 Then
                Call RecordBuffer(1, Me.vsList)
                If mlngModi = 0 Then
                    mlngModi = lng��ˮ��                '�޸�״̬
                    Call mfrmMain.rptRecord_SelectionChanged
                Else
                    MsgBox "����һ�ŵ��������޸ģ�����ͬʱ�޸Ķ��ŵ��ݣ�", vbInformation, gstrSysName
                End If
            End If
        Case conMenu_Edit_Transf_Save
            '����
            'lng��ˮ�� = Val(vsMain.TextMatrix(vsMain.Row, col_��ˮ��))
            If mlngModi <> 0 Then
                '###
                Call SaveToExecRecord(mlngModi, lngErrNo)
                mlngModi = 0
                If lngErrNo = 0 Then
                    'Call mfrmMain.rptRecord_SelectionChanged
                    Call mfrmMain.rptPati_SelectionChanged
                Else
                    Call mfrmMain.rptRecord_SelectionChanged
                    Call RecordBuffer(2, Me.vsList)
                End If
            End If
        Case conMenu_File_BillPrint, conMenu_File_BillPrintEx
            '���ݴ�ӡ
            lng��ˮ�� = mfrmMain.Get��ˮ��
            If lng��ˮ�� <> 0 Then
                If Control.ID = conMenu_File_BillPrintEx Then
                    Select Case Val(Control.Caption)
                    Case 1    '��Һƿǩ
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_4", mfrmMain, "�ӵ���ˮ��=" & lng��ˮ��)
                    Case 2    '���
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_5", mfrmMain, "�ӵ���ˮ��=" & lng��ˮ��)
                    End Select
                Else
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_" & mlngType, mfrmMain, "�ӵ���ˮ��=" & lng��ˮ��, 2)
                End If
            End If
        Case conMenu_File_BillPrintView
            '����Ԥ��
            lng��ˮ�� = mfrmMain.Get��ˮ��
            If lng��ˮ�� <> 0 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_" & mlngType, mfrmMain, "�ӵ���ˮ��=" & lng��ˮ��, 1)
            End If
'        Case conMenu_Edit_Transf_Positive
'            '����(-)
'            If mlngModi = 0 Then
'                lng��ˮ�� = mfrmMain.Get��ˮ��
'                If lng��ˮ�� <> 0 And mstrGroupKey <> "" Then
'                    strGroupKey = mstrGroupKey
'                    Call mfrmMain.ExecuteTest(CStr(lng��ˮ��), strGroupKey, "(-)")
'                    Call mfrmMain.rptRecord_SelectionChanged
'                End If
'            End If
'        Case conMenu_Edit_Transf_Negative
'            '����(+)
'            If mlngModi = 0 Then
'                lng��ˮ�� = mfrmMain.Get��ˮ��
'                If lng��ˮ�� <> 0 And mstrGroupKey <> "" Then
'                    strGroupKey = mstrGroupKey
'                    Call mfrmMain.ExecuteTest(CStr(lng��ˮ��), strGroupKey, "(+)")
'                    Call mfrmMain.rptRecord_SelectionChanged
'                End If
'            End If
        Case conMenu_Edit_Test
            'Ƥ�Խ��
            If mlngModi = 0 Then
                lng��ˮ�� = mfrmMain.Get��ˮ��
                If lng��ˮ�� <> 0 And mstrGroupKey <> "" Then
                    strGroupKey = mstrGroupKey
                    Call mfrmMain.ExecuteTest(CStr(lng��ˮ��), strGroupKey)
                    Call mfrmMain.ˢ��
                End If
            End If
        Case conMenu_Edit_Transf_Cancle
            'ȡ��
            If mlngModi <> 0 Then
                If MsgBox("�Ƿ�ȡ�����޸ĵ����ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    mlngModi = 0
                    Call mfrmMain.rptRecord_SelectionChanged
                End If
            End If
         Case conMenu_Manage_Undone
            '�������
            If mlngModi = 0 Then
                lng��ˮ�� = mfrmMain.Get��ˮ��
                lngDeptID = mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex)
                If lng��ˮ�� <> 0 And mstrGroupKey <> "" Then
                    strGroupKey = mstrGroupKey
                    
                    If mfrmMain.mobjRecord.Item(CStr(lng��ˮ��)).Item(strGroupKey).ִ��״̬ = 1 Then
                        '����ɵģ��ȸ�״̬Ϊִ����
                        If mfrmMain.mobjRecord.Item(CStr(lng��ˮ��)).ExecCanle(strGroupKey, mfrmMain.mblnƤ����֤, lngDeptID, Me) = False Then
                            Exit Sub
                        End If
                    End If
                    If (mfrmMain.mobjRecord.Item(CStr(lng��ˮ��)).Item(strGroupKey).�˶��� = "" And Mid(gstrҽ���˶�, 2, 1) <> "1") Or (mfrmMain.mobjRecord.Item(CStr(lng��ˮ��)).Item(strGroupKey).�˶��� = "" And Mid(gstrҽ���˶�, 2, 1) = "1" And mfrmMain.mobjRecord.Item(CStr(lng��ˮ��)).Item(strGroupKey).ִ��״̬ <> 1) Then
                        Call mfrmMain.ExecStart(CStr(lng��ˮ��), strGroupKey, True)
                    End If
                    'Call mfrmMain.rptRecord_SelectionChanged
                    Call mfrmMain.ˢ��
                End If
            End If
        Case conMenu_Manage_Complete
            '��� 2012��09��10 ��Ϊ��ʼ���ܣ���дÿ����Һҽ���Ŀ�ʼʱ�䣬��ʼ��,�������һ�Σ���Ҫ��ԭ������ɹ���
            'ʣ�������Ϊ0������ִ����ɹ��� ,objGroup.�������� - objGroup.��ִ������
            '
            If mlngModi = 0 Then
                lng��ˮ�� = mfrmMain.Get��ˮ��
                If lng��ˮ�� <> 0 And mstrGroupKey <> "" Then
                    strGroupKey = mstrGroupKey
                    If mfrmMain.ExecStart(CStr(lng��ˮ��), strGroupKey) Then
                    
                        If Val(vsList.TextMatrix(vsList.Row, col_ʣ�����)) = 0 Then
                            
                            Call mfrmMain.ExecComplt(CStr(lng��ˮ��), strGroupKey)
                            Call mfrmMain.ˢ��
                        Else
                            '----- �ĳ��޸� ����ҽ��ִ�� ��Ŀ�ʼʱ��
                            'MsgBox "����Ŀ����ʣ��ִ�д����������ܱ��Ϊ��ɣ�", vbInformation, gstrSysName
    
                            Call mfrmMain.ˢ��
                        End If
                    End If
                End If
            End If
'�ܾ���ȡ���ܾ�,�ŵ��ӵ��д���
'        Case conMenu_Manage_Refuse
'            '�ܾ�ִ��
'            If mlngModi = 0 Then
'                lng��ˮ�� = mfrmMain.Get��ˮ��
'                If lng��ˮ�� <> 0 And mstrGroupKey <> "" Then
'                    Call mfrmMain.mobjRecord.Item(CStr(��ˮ��)).FuncExecRefuse(mstrGroupKey)
'                    Call mfrmMain.rptRecord_SelectionChanged
'                End If
'            End If
'
'        Case conMenu_Manage_ReGet
'            'ȡ���ܾ�
'            If mlngModi = 0 Then
'                lng��ˮ�� = mfrmMain.Get��ˮ��
'                If lng��ˮ�� <> 0 And mstrGroupKey <> "" Then
'                    Call mfrmMain.mobjRecord.Item(CStr(��ˮ��)).FuncExecRestore(mstrGroupKey)
'                    Call mfrmMain.rptRecord_SelectionChanged
'                End If
'            End If
        Case conMenu_Manage_ThingAudit '�˶�
            If mlngModi = 0 Then
                lng��ˮ�� = mfrmMain.Get��ˮ��
                strGroupKey = mstrGroupKey
                If lng��ˮ�� <> 0 And strGroupKey <> "" Then
                    If mfrmMain.FuncThingAudit(CStr(lng��ˮ��), strGroupKey) Then
                        mfrmMain.ˢ��
                    End If
                End If
            End If
        Case conMenu_Manage_ThingDelAudit 'ȡ���˶�
            If mlngModi = 0 Then
                lng��ˮ�� = mfrmMain.Get��ˮ��
                strGroupKey = mstrGroupKey
                If lng��ˮ�� <> 0 And strGroupKey <> "" Then
                    If mfrmMain.FuncThingDelAudit(CStr(lng��ˮ��), strGroupKey) Then
                        mfrmMain.ˢ��
                    End If
                End If
            End If
        Case Else
            If Not tbcKernel.Selected Is Nothing Then
                If tbcKernel.Selected.Tag = "��Ŀ����" Then
                    Call mclsExpenses.zlExecuteCommandBars(Control)
                End If
            End If
    End Select
    
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    '����������ñ�����ĵ����˵�
    Call mclsExpenses.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    '��������ñ������״̬����
    Dim blnVisable As Boolean, blnAllEnable As Boolean, blnOneEnable As Boolean, bln״̬ As Boolean
    Dim intItem As Integer
    
    blnVisable = (mlngType = 3)
    'blnAllEnable = (mlng��ˮ�� > 0 And InStr(mstrGroupKey, "_") <= 0 And mlngModi = 0) '����ˮ��
    blnAllEnable = (mlng��ˮ�� > 0 And mlngModi = 0)  '����ˮ��
    blnOneEnable = (mlng��ˮ�� > 0 And InStr(mstrGroupKey, "_") <> 0 And mlngModi = 0) '��ִ����Ŀ
    
    If mfrmMain.mobjRecord Is Nothing Then
        blnAllEnable = False
        blnOneEnable = False
    End If
    
    Select Case Control.ID
        Case conMenu_Manage_ThingDel, conMenu_File_BillPrint, conMenu_File_BillPrintView, conMenu_File_BillPrintEx
            '�����ӵ�
            Control.Enabled = blnAllEnable

            If Control.Enabled Then
                 
                For intItem = 1 To mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Count
                    If mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(intItem).ִ��״̬ = 1 Then
                        Control.Enabled = False
                        Exit For
                    End If
                    If blnVisable And Mid(gstrҽ���˶�, 2, 1) = "1" And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).�˶��� <> "" Then
                        Control.Enabled = False
                        Exit For
                    End If
                Next
     
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-�˺�"
        Case conMenu_Edit_Transf_Modify
            '�޸�
            Control.Enabled = mlngModi = 0 And mlng��ˮ�� > 0
            If InStr(mstrGroupKey, "_") <= 0 Then Control.Enabled = False
            If Control.Enabled Then
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ <> 1
                    If blnVisable And Mid(gstrҽ���˶�, 2, 1) = "1" Then
                        If Control.Enabled Then Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).�˶��� = ""
                    End If
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-�˺�"
        Case conMenu_Edit_Transf_Save
            '����
            Control.Enabled = mlngModi <> 0 And mlng��ˮ�� > 0
            If InStr(mstrGroupKey, "_") <= 0 Then Control.Enabled = False
            If Control.Enabled Then
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ <> 1
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-�˺�"
        Case conMenu_Edit_Transf_Cancle
            'ȡ��
            Control.Enabled = mlngModi <> 0 And mlng��ˮ�� > 0
            If InStr(mstrGroupKey, "_") <= 0 Then Control.Enabled = False
            If Control.Enabled Then
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ <> 1
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-�˺�"
'        Case conMenu_Manage_Refuse
'            '�ܾ�
'
'        Case conMenu_Manage_ReGet
'            'ȡ���ܾ�
        Case conMenu_Manage_Complete
            '���--  �ĳ��˿�ʼ
            Control.Visible = Not blnVisable
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ <> 1
                    
                    If Control.Enabled Then
                        Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ���� = ""
                    End If
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-�˺�"
        Case conMenu_Manage_Undone
            'ȡ�����
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And (mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ���� <> "" _
                      Or mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ = 1)
                    If blnVisable And Mid(gstrҽ���˶�, 2, 1) = "1" Then
                         Control.Enabled = (mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).�˶��� <> "" And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ = 1) Or (mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).�˶��� = "" And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ = 3)
                    End If
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-�˺�"

'        Case conMenu_Edit_Transf_Negative
'            '����
'            Control.Visible = blnVisable
'            If InStr(mstrGroupKey, "_") <= 0 Then
'                Control.Enabled = False
'            Else
'                If Not mfrmMain.mobjRecord Is Nothing Then
'                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ <> 1
'                Else
'                    Control.Enabled = False
'                End If
'            End If
'            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-�˺�"
'        Case conMenu_Edit_Transf_Positive
'            '����
'            Control.Visible = blnVisable
'            If InStr(mstrGroupKey, "_") <= 0 Then
'                Control.Enabled = False
'            Else
'                If Not mfrmMain.mobjRecord Is Nothing Then
'                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ <> 1
'                Else
'                    Control.Enabled = False
'                End If
'            End If
'            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-�˺�"
        Case conMenu_Edit_Test
            'Ƥ�Խ��
            Control.Visible = blnVisable
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ <> 1
                    If blnVisable And Mid(gstrҽ���˶�, 2, 1) = "1" Then
                        If Control.Enabled Then Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).�˶��� <> ""
                    End If
                Else
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Manage_ThingAudit
            'ҽ���˶�
            Control.Visible = blnVisable
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ <> 1 And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).�˶��� = ""
                Else
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Manage_ThingDelAudit
            'ҽ��ȡ���˶�
            Control.Visible = blnVisable
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).�˶��� <> "" And mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(mstrGroupKey).ִ��״̬ <> 1
                Else
                    Control.Enabled = False
                End If
            End If
        Case Else
            If Not tbcKernel.Selected Is Nothing Then
                If tbcKernel.Selected.Tag = "��Ŀ����" Then
                    Call mclsExpenses.zlUpdateCommandBars(Control)
                End If
            End If
    End Select

End Sub

Public Sub zlRefresh(ByVal objRecord As ExecRecord, ByVal objPati As cPatient)
    '������Ҫ�󱾴���ˢ��
    Call mfrmMain.ShowReport
    Call ShowVsList(mfrmMain.Get��ˮ��)
    Call KernalRefresh
    
    'ִ����
    mstrִ���� = ""
    
    Dim ObjOutNurse As New OutNurses, objNurs As OutNurse
    
    If mstrִ���� = "" And mfrmMain.cboDept.ListIndex >= 0 Then
        ObjOutNurse.getOutNurse (mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex))
        For Each objNurs In ObjOutNurse
            mstrִ���� = mstrִ���� & "|" & objNurs.����
        Next
        If Mid(mstrִ����, 1, 1) = "|" Then mstrִ���� = Mid(mstrִ����, 2)
    End If
    '��ǰ���˵�״̬
    If Not objPati Is Nothing Then
        mstrPatiStat = objPati.�Ŷ�״̬
    Else
        mstrPatiStat = ""
    End If
End Sub

Public Sub KernalRefresh()
    Dim strInfo As String

    If mstrGroupKey = "" Then
        Call mclsExpenses.zlRefresh(0, "")
    Else
        '����ID,ҽ��ID,���ͺ�
        strInfo = Trim(Split(mstrGroupKey, "_")(0)) & ":" & Split(mstrGroupKey, "_")(1)
'        On Error Resume Next
'        Call mclsExpenses.zlRefresh(mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex), _
                                    Val(Split(mstrGroupKey, "_")(0)), Val(Split(mstrGroupKey, "_")(1)), False)
        Call mclsExpenses.zlRefresh(mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex), _
                                    strInfo)
    End If
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object)
    '������Ҫ���ʼ���������ϵĲ˵�
    
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '������Ŀ�Ĳ˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set mcbsMain = cbsMain
    
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "ִ��(&E)", objMenu.Index + 1, False)
    End If
    
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "�����ӵ�(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Modify, "�޸�(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��(&C)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_BillPrintEx, "1-��ӡ��Һƿǩ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_BillPrintEx, "2-��ӡ�����ǩ")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_BillPrint, "�ش򵥾�(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_BillPrintView, "����Ԥ��(&V)")
        
'        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "�ܾ�ִ��(&R)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "ȡ���ܾ�(&G)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "��ʼִ��(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "����ִ��(&U)")
        
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Negative, "����(&+)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Positive, "����(&-)")
        If Mid(gstrҽ���˶�, 2, 1) = "1" Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "�˶�"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit, "ȡ���˶�")
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Test, "Ƥ�Խ��(&H)"): objControl.BeginGroup = True

    End With
    
    '����˵�:���������û��,���ڲ鿴�˵�ǰ��
    '-----------------------------------------------------
    '����վ����˵��Զ���ʾ��������Թ���վ��ģ���ͳһ����
    '���⼸�ű�����ҽ������ģ���еģ���Ҫ�ڸ�ģ���е�������
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
'    If objMenu Is Nothing Then
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)", objMenu.Index, False)
'        objMenu.ID = conMenu_ReportPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
'    End If
'    With objMenu.CommandBar.Controls
'        '���������ǰ��,�������
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Reprint, "�ش򵥾�(&R)", 1)
'    End With
    
    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '�����ǰ������һ��Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Modify, "�޸�", objControl.Index + 1): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "ִ��", objControl.Index + 1): objControl.BeginGroup = True
        objControl.ToolTipText = "��ʼִ��"
    End With
    Set mfrmMain = frmParent
    cbsSub.ActiveMenuBar.Visible = False
    '--��Ŀ����
    Call mclsExpenses.zlDefCommandBars(frmParent, cbsMain)
    Set mcbsMain.Icons = zlCommFun.GetPubIcons


End Sub

Public Sub ShowVsList(ByVal lng��ˮ�� As Long)
    
    Dim objGroup As Group, objBIll As Bill, lng��� As Long, strListHead As String, str״̬ As String
    Dim date��ʼʱ�� As Date, lng�ѿ�ʼ As Long
    
    If lng��ˮ�� = 0 Then
        mlngType = 0
    Else
        mlngType = Mid(mfrmMain.mobjRecord.Item(CStr(lng��ˮ��)).ִ�з���, 1, 1)
    End If
    
    Select Case mlngType
    Case 1
        '��Һ
        strListHead = "״̬,450,4;���,0,4;˳��,450,4;����,2300,1;����,600,7;��λ,0,4;���,700,7;ִ��Ƶ��,900,1;�÷�,900,1;Ƥ�Խ��,0,1;��Һ��,700,7;����,450,7;����(ml),450,7;ʱ��(��),450,4;����,1300,1;ִ����,675,1;�˶���,0,1;ʣ�����,450,7;��ע,500,1;billKey,0,1;GroupKey,0,1;�޸ı�־,0,1"
    Case 2
        'ע��
        strListHead = "״̬,450,4;���,0,4;˳��,0,4;����,2300,1;����,600,7;��λ,0,4;���,700,7;ִ��Ƶ��,900,1;�÷�,900,1;Ƥ�Խ��,0,1;ע���,700,7;����,0,7;����(ml),0,7;ʱ��(��),0,4;����,0,1;ִ����,675,1;�˶���,0,1;ʣ�����,450,7;��ע,500,1;billKey,0,1;GroupKey,0,1;�޸ı�־,0,1"
    Case 3
        'Ƥ��
        If Mid(gstrҽ���˶�, 2, 1) = "1" Then
            strListHead = "״̬,450,4;���,0,4;˳��,0,4;����,2300,1;����,0,7;��λ,0,4;���,0,7;ִ��Ƶ��,900,1;�÷�,1800,1;Ƥ�Խ��,450,1;Ƥ�Է�,700,7;����,0,7;����(ml),0,7;ʱ��(��),0,4;����,1300,1;ִ����,675,1;�˶���,675,1;ʣ�����,450,7;��ע,500,1;billKey,0,1;GroupKey,0,1;�޸ı�־,0,1"
        Else
            strListHead = "״̬,450,4;���,0,4;˳��,0,4;����,2300,1;����,0,7;��λ,0,4;���,0,7;ִ��Ƶ��,900,1;�÷�,1800,1;Ƥ�Խ��,450,1;Ƥ�Է�,700,7;����,0,7;����(ml),0,7;ʱ��(��),0,4;����,1300,1;ִ����,675,1;�˶���,0,1;ʣ�����,450,7;��ע,500,1;billKey,0,1;GroupKey,0,1;�޸ı�־,0,1"
        End If
    Case Else
        '����
        '1 ����� 4 ���� 7 �Ҷ���
        strListHead = "״̬,450,4;���,0,4;˳��,0,4;����,2300,1;����,600,7;��λ,0,4;���,0,7;ִ��Ƶ��,900,1;�÷�,900,1;Ƥ�Խ��,0,1;���Ʒ�,700,7;����,0,7;����(ml),0,7;ʱ��(��),0,4;����,0,1;ִ����,675,1;�˶���,0,1;ʣ�����,450,7;��ע,500,1;billKey,0,1;GroupKey,0,1;�޸ı�־,0,1"
    End Select
    Call SetVsFlexGridHead(strListHead, vsList)
    If lng��ˮ�� = 0 Then Exit Sub
    If mfrmMain.mobjRecord Is Nothing Then Exit Sub
    'vsList.Redraw = False
    date��ʼʱ�� = mfrmMain.mobjRecord.Item(CStr(lng��ˮ��)).ִ��ʱ��
    lng�ѿ�ʼ = DateDiff("n", date��ʼʱ��, zlDatabase.Currentdate)
    
    For Each objGroup In mfrmMain.mobjRecord.Item(CStr(lng��ˮ��))
        lng��� = 0
        date��ʼʱ�� = date��ʼʱ�� + (objGroup.��ʱ / 24 / 60)
        
        
        With vsList
            For Each objBIll In objGroup.BillsItem(objGroup.ִ��ҽ��ID & "_" & objGroup.���ͺ�)
                lng��� = lng��� + 1
                '0-δִ��;1-��ȫִ��;2-�ܾ�ִ��;3-����ִ��
                Select Case objGroup.ִ��״̬
                    Case 0
                        str״̬ = ""
                    Case 1
                        str״̬ = "���"
                        'Set .Cell(flexcpPicture, .Rows - 1, col_������) = img���.Picture
                        .TextMatrix(.Rows - 1, col_������) = "��"
                    Case 2
                        str״̬ = "�ܾ�"
                        'Set .Cell(flexcpPicture, .Rows - 1, col_������) = img�ܾ�.Picture
                        .TextMatrix(.Rows - 1, col_������) = "��"
                    Case Else
                        str״̬ = "ִ����"
                        .TextMatrix(.Rows - 1, col_������) = "��"
                        'Set .Cell(flexcpPicture, .Rows - 1, col_������) = img��ִ��.Picture
                End Select
                '.TextMatrix(.Rows - 1, col_������) = str״̬
                
                .TextMatrix(.Rows - 1, col_���) = lng���
                .TextMatrix(.Rows - 1, col_˳��) = Val(objGroup.���)
                .TextMatrix(.Rows - 1, col_ҽ������) = objBIll.ҽ������
                .TextMatrix(.Rows - 1, col_����) = IIf(Mid(objBIll.����, 1, 1) = ".", "0" & objBIll.���� & objBIll.��λ, objBIll.���� & objBIll.��λ)
                .TextMatrix(.Rows - 1, col_��λ) = objBIll.��λ
                .TextMatrix(.Rows - 1, col_���) = IIf(Format(objBIll.���, "0.00") = 0, "", Format(objBIll.���, "0.00"))
                If objBIll.��ϸ�Ʒ�״̬ = -1 Then
                    .TextMatrix(.Rows - 1, col_���) = "���Ʒ�"
                ElseIf objBIll.��ϸ�Ʒ�״̬ = -2 Then
                    If objBIll.��� = 0 Then .TextMatrix(.Rows - 1, col_���) = "�����"
                ElseIf objBIll.��ϸ�Ʒ�״̬ = -3 Then
                    .TextMatrix(.Rows - 1, col_���) = "���˷�"
                End If
                .TextMatrix(.Rows - 1, col_ִ��Ƶ��) = objGroup.ִ��Ƶ��
                .TextMatrix(.Rows - 1, col_�÷�) = objGroup.�÷�
                .TextMatrix(.Rows - 1, col_Ƥ�Խ��) = objGroup.Ƥ�Խ��
                .TextMatrix(.Rows - 1, col_�շѽ��) = IIf(Format(objGroup.�շѽ��, "0.00") = 0, "", Format(objGroup.�շѽ��, "0.00"))
                If objGroup.�Ʒ�״̬ = -1 Then
                    .TextMatrix(.Rows - 1, col_�շѽ��) = "���Ʒ�"
                ElseIf objGroup.�Ʒ�״̬ = -2 Then
                    If objGroup.�շѽ�� = 0 Then .TextMatrix(.Rows - 1, col_�շѽ��) = "�����"
                ElseIf objGroup.�Ʒ�״̬ = -3 Then
                    .TextMatrix(.Rows - 1, col_�շѽ��) = "���˷�"
                End If
                .TextMatrix(.Rows - 1, col_����) = objGroup.����
                .TextMatrix(.Rows - 1, col_����) = objGroup.Һ����
                .TextMatrix(.Rows - 1, col_ʱ��) = objGroup.��ʱ
                
                '.TextMatrix(.Rows - 1, col_����) = Format(date��ʼʱ��, "MM-dd hh:mm")
                
                If lng�ѿ�ʼ >= objGroup.��ʱ Then
                    lng�ѿ�ʼ = lng�ѿ�ʼ - objGroup.��ʱ
                    .Cell(flexcpData, .Rows - 1, col_����) = 100
                    '.Cell(flexcpFloodPercent, .Rows - 1, col_����) = 100
                    '.Cell(flexcpFloodColor, .Rows - 1, col_����) = RGB(215, 215, 235)
                    
                Else
                    If lng�ѿ�ʼ >= 0 Then
                        .Cell(flexcpData, .Rows - 1, col_����) = (lng�ѿ�ʼ / objGroup.��ʱ) * 100
                        '.Cell(flexcpFloodPercent, .Rows - 1, col_����) = (lng�ѿ�ʼ / objGroup.��ʱ) * 100
                        '.Cell(flexcpFloodColor, .Rows - 1, col_����) = RGB(215, 215, 235)
                        lng�ѿ�ʼ = lng�ѿ�ʼ - objGroup.��ʱ
                    End If
                End If
                .TextMatrix(.Rows - 1, col_˵��) = IIf(objBIll.ҽ������ = "", "��", objBIll.ҽ������)
                .TextMatrix(.Rows - 1, col_ִ����) = IIf(objGroup.ִ���� = "", "��", objGroup.ִ����)
                .TextMatrix(.Rows - 1, col_�˶���) = IIf(objGroup.�˶��� = "", "��", objGroup.�˶���)
                .TextMatrix(.Rows - 1, col_ʣ�����) = objGroup.�������� - objGroup.��ִ������ - objGroup.��������
                .TextMatrix(.Rows - 1, col_BillKey) = objGroup.ִ��ҽ��ID & "_" & objBIll.ҽ��ID
                .TextMatrix(.Rows - 1, col_groupkey) = objGroup.ִ��ҽ��ID & "_" & objGroup.���ͺ�
                
                '�ֶ���ɫ��Ƥ���ࣨ���ԣ���ɫ�����ԣ���ɫ������Ƥ�����ɫ��
                If InStr(objGroup.Ƥ�Խ��, "(-)") > 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HFF0000
                ElseIf InStr(objGroup.Ƥ�Խ��, "(+)") > 0 Or InStr(objGroup.Ƥ�Խ��, "(++)") > 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0&
                End If
                                
                .Rows = .Rows + 1
                '.TextMatrix(.Rows - 1, col_groupkey) = "Ҫ����"
            Next
            'ÿ��ҩ֮���һ����,������ͬ���ݱ��ϲ�Ϊһ����Ԫ��
            '.Rows = .Rows + 1
'            If .TextMatrix(.Rows - 2, col_groupkey) = "Ҫ����" Then
'                .RowHidden(.Rows - 2) = True
'            End If
        End With
    Next
    If vsList.Rows > 2 Then
        vsList.RemoveItem vsList.Rows - 1
    End If
    'vsMain.Redraw = True
    With vsList
'        .MergeCells = flexMergeRestrictColumns
'        .MergeCol(col_������) = True
'        .MergeCol(col_ִ��Ƶ��) = True
'        .MergeCol(col_�÷�) = True
'        .MergeCol(col_����) = True
'        .MergeCol(col_����) = True
'        .MergeCol(col_ʱ��) = True
'        .MergeCol(col_ʣ�����) = True
'        .MergeCol(col_ִ����) = True
'        .MergeCol(col_����) = True
'        .MergeCol(col_˵��) = True
        .AutoSize 1, col_groupkey
        .RowHeight(0) = 500
        '.BackColorSel = .BackColor
        '.ForeColorSel = .ForeColor
        '.SelectionMode = flexSelectionFree
    End With
    
    vsList.ColDataType(col_����) = flexDTLong
    vsList.ColDataType(col_����) = flexDTLong
    
    If mlngModi <> 0 Then
        '�޸�״̬
        vsList.Cell(flexcpBackColor, 1, col_����, vsList.Rows - 1, col_����) = VsModiBackColor
        vsList.Cell(flexcpBackColor, 1, col_����, vsList.Rows - 1, col_����) = VsModiBackColor
        vsList.Cell(flexcpBackColor, 1, col_˵��, vsList.Rows - 1, col_˵��) = VsModiBackColor
        vsList.Cell(flexcpBackColor, 1, col_ִ����, vsList.Rows - 1, col_ִ����) = VsModiBackColor
        
        If mstrִ���� <> "" Then
            vsList.ColComboList(col_ִ����) = mstrִ����
        End If
        vsList.Editable = flexEDKbdMouse
        
        'vsMain.Cell(flexcpForeColor, vsMain.Row, 0, vsMain.Row, vsMain.Cols - 1) = vbRed
        mblnUpdate = True
    Else
        '�鿴״̬
        
        vsList.Cell(flexcpBackColor, 1, col_����, vsList.Rows - 1, col_����) = vsList.BackColor
        vsList.Cell(flexcpBackColor, 1, col_����, vsList.Rows - 1, col_����) = vsList.BackColor
        vsList.Cell(flexcpBackColor, 1, col_˵��, vsList.Rows - 1, col_˵��) = vsList.BackColor
        vsList.Cell(flexcpBackColor, 1, col_ִ����, vsList.Rows - 1, col_ִ����) = vsList.BackColor
        vsList.Editable = flexEDNone
        'vsMain.Cell(flexcpForeColor, vsMain.Row, 0, vsMain.Row, vsMain.Cols - 1) = vsMain.ForeColor
        mblnUpdate = False
    End If
    vsList_RowColChange
End Sub


Private Sub tbcKernel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsList.AutoSize 1, col_groupkey
    vsList.RowHeight(0) = 500
End Sub

Private Sub vsList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If col_groupkey < vsList.Cols Then
        If mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).Item(vsList.TextMatrix(Row, col_groupkey)).ִ��״̬ = 1 Then
            '��ɵļ�¼�����޸�
            Cancel = True
            Exit Sub
        End If
    End If
    If InStr("," & col_���� & "," & col_���� & "," & col_˵�� & "," & col_ִ���� & ",", _
             "," & Col & ",") <= 0 Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub vsList_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim LeftCol As Long, RightCol As Long, topRow As Long, BottomRow As Long
    
    If Col = col_���� Then
        If Val(vsList.TextMatrix(Row, col_���)) = 1 And Trim(vsList.TextMatrix(Row, col_ִ����)) <> "" Then
            With vsList
                Call vfgDrawProgress(vsList, Row, Col, hDC, Left, Top, Right, Bottom, .Cell(flexcpData, Row, Col))
            End With
        End If
    End If
    
    
    If Not MergeRow(Row, topRow, BottomRow) Then Exit Sub '�Ǻϲ���,�˳�
    If topRow = BottomRow Then Exit Sub
    
    LeftCol = col_˳��: RightCol = col_˳��
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
    LeftCol = col_������: RightCol = col_������
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
    LeftCol = col_ִ��Ƶ��: RightCol = col_�÷�
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)

    LeftCol = col_����: RightCol = col_ʣ�����
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
    LeftCol = col_˵��: RightCol = col_˵��
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
    LeftCol = col_�շѽ��: RightCol = col_�շѽ��
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
End Sub

Private Function MergeRow(ByVal Row As Long, topRow, BottomRow As Long) As Boolean
    '�Ƿ�ϲ���
    Dim strGroupKey As String, lngRow As Long
    With vsList
        If .Cols < col_groupkey Then Exit Function
        strGroupKey = .TextMatrix(Row, col_groupkey)
        topRow = Row: BottomRow = Row
        For lngRow = Row To 0 Step -1
            If .TextMatrix(lngRow, col_groupkey) <> strGroupKey Then
                topRow = lngRow + 1
                Exit For
            Else
                topRow = lngRow
            End If
        Next
        
        For lngRow = Row To .Rows - 1
            If .TextMatrix(lngRow, col_groupkey) <> strGroupKey Then
                BottomRow = lngRow - 1
                Exit For
            Else
                BottomRow = lngRow
            End If
        Next
    End With

    If topRow > 0 And BottomRow > 0 Then MergeRow = True
End Function

Private Sub vsList_EnterCell()
    If mlngModi = 1 Then
        If vsList.Col = col_���� Or vsList.Col = col_���� Or vsList.Col = col_˵�� Or vsList.Col = col_ִ���� Then
            Call vsList.CellBorder(vsList.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End If
End Sub

Private Sub vsList_GotFocus()
    If col_groupkey < vsList.Cols Then
        mstrGroupKey = Trim(vsList.TextMatrix(vsList.Row, col_groupkey))
    Else
        mstrGroupKey = ""
    End If
End Sub

Private Sub vsList_LeaveCell()
    If mlngModi = 1 Then
        If vsList.Col = col_���� Or vsList.Col = col_���� Or vsList.Col = col_˵�� Or vsList.Col = col_ִ���� Then
            On Error Resume Next
            Call vsList.CellBorder(vsList.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End If
End Sub

Private Sub vsList_LostFocus()
    mstrGroupKey = ""
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_ManagePopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub vsList_RowColChange()
    If col_groupkey < vsList.Cols Then
        
        If mstrGroupKey <> Trim(vsList.TextMatrix(vsList.Row, col_groupkey)) Then
            mstrGroupKey = Trim(vsList.TextMatrix(vsList.Row, col_groupkey))
            Call KernalRefresh
        End If
    Else
        mstrGroupKey = ""
    End If
    
    '����ѡ�еĵ�ǰ������ǰ��ɫ��ԭ����ǰ��ɫ��ͬ
    If vsList.Row > 0 Then
        vsList.ForeColorSel = vsList.Cell(flexcpForeColor, vsList.Row, 0)
    End If
End Sub

Private Sub SaveToExecRecord(ByVal str��ˮ�� As String, Optional ByRef lngErrNo_Out As Long)
    '�����޸ĵ���Ϣ
    Dim iRow As Integer, strGroupKey As String, strBillKey As String, intִ��״̬ As Integer, blnLoad��� As Boolean
    Dim lngDeptID As Long
    Dim cnNew As ADODB.Connection, strUserName As String
    Dim lngErrNo As Long
    
    If str��ˮ�� = "" Then Exit Sub
    '11471 ����޸�ʱ��û�а��س�,�򲻱����޸ĵ����ݡ�
    If vsList.Col < vsList.Cols - 1 Then
        vsList.Select vsList.Row, vsList.Col + 1
    Else
        vsList.Select vsList.Row, vsList.Col - 1
    End If
    
    For iRow = 1 To vsList.Rows - 1
        
        If vsList.TextMatrix(iRow, col_�޸ı�־) = "Update" Then
            strGroupKey = vsList.TextMatrix(iRow, col_groupkey)
            
            If vsList.TextMatrix(iRow, col_���) = 1 Then
                strBillKey = vsList.TextMatrix(iRow, col_BillKey)
                mfrmMain.mobjRecord.Item(str��ˮ��).Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).���� = Val(vsList.TextMatrix(iRow, col_����))
                mfrmMain.mobjRecord.Item(str��ˮ��).Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).ʱ�� = Val(vsList.TextMatrix(iRow, col_ʱ��))
            End If
            
            mfrmMain.mobjRecord.Item(str��ˮ��).Item(strGroupKey).���� = Val(vsList.TextMatrix(iRow, col_����))
            mfrmMain.mobjRecord.Item(str��ˮ��).Item(strGroupKey).˵�� = Replace(vsList.TextMatrix(iRow, col_˵��), "��", "")
            mfrmMain.mobjRecord.Item(str��ˮ��).Item(strGroupKey).ִ���� = Replace(vsList.TextMatrix(iRow, col_ִ����), "��", "")
            
            lngDeptID = mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex)
            Call mfrmMain.mobjRecord.Item(str��ˮ��).Update(str��ˮ��, strGroupKey, lngDeptID, lngErrNo)
            
            If lngErrNo <> 0 Then
                lngErrNo_Out = lngErrNo
                Exit Sub
            End If
            
            If str��ˮ�� = mfrmMain.Get��ˮ�� Then
                mfrmMain.rptRecord.SelectedRows(0).Record(rptCOL_��ʱ).Value = mfrmMain.mobjRecord.Item(str��ˮ��).�ܺ�ʱ
                mfrmMain.rptRecord.SelectedRows(0).Record(rptCOL_��ʱ).Caption = mfrmMain.mobjRecord.Item(str��ˮ��).�ܺ�ʱ
                mfrmMain.rptRecord.Populate
            End If
        End If
    Next

End Sub

Private Sub vsList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim blnExit As Boolean, i As Integer
    
    Select Case Col
    Case col_������
        If Val(vsList.TextMatrix(Row, col_ʣ�����)) <> 0 And vsList.EditText = "���" Then
            Cancel = True
            MsgBox "ʣ�������Ϊ0�����ڻ�������ɣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    Case col_����
        i = 0: blnExit = False
        If Val(Trim(vsList.TextMatrix(Row + i, col_���))) = 0 Then Exit Sub
        If Val(vsList.EditText) < 0 Or Val(vsList.EditText) > 10000 Then
            Cancel = True: Exit Sub
        End If
        Do While blnExit = False
            If blnExit = False Then
                If Row + i >= vsList.Rows Then
                    blnExit = True
            
                ElseIf (Val(vsList.TextMatrix(Row + i, col_���)) = 1 Or Val(vsList.TextMatrix(Row + i, col_���)) = 0) And i > 0 Then
                    blnExit = True
                Else
                    vsList.TextMatrix(Row + i, col_ʱ��) = CacleTransTime(Val(vsList.EditText), _
                                                                     mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).��ϵ��, _
                                                                     Val(vsList.TextMatrix(Row, col_����)))
                    i = i + 1
                End If
            End If
        Loop
    Case col_����
        If Val(vsList.EditText) < 10 Or Val(vsList.EditText) > 100 Then
             Cancel = True: Exit Sub
        End If
        i = 0: blnExit = False
        If Val(Trim(vsList.TextMatrix(Row + i, col_���))) = 0 Then Exit Sub
        Do While blnExit = False
            If blnExit = False Then
                If Row + i >= vsList.Rows Then
                    blnExit = True
                ElseIf (Val(vsList.TextMatrix(Row + i, col_���)) = 1 Or Val(vsList.TextMatrix(Row + i, col_���)) = 0) And i > 0 Then
                    blnExit = True
                Else
                    
                    vsList.TextMatrix(Row + i, col_ʱ��) = CacleTransTime(Val(vsList.TextMatrix(Row, col_����)), _
                                                                      mfrmMain.mobjRecord.Item(CStr(mlng��ˮ��)).��ϵ��, _
                                                                     Val(vsList.EditText))
                    i = i + 1
                End If
            End If
        Loop
    Case col_Ƥ�Խ��
        If InStr(",(+),(-),����", vsList.TextMatrix(Row, col_Ƥ�Խ��)) > 0 Then
            Cancel = True
            MsgBox "����д����ļ�¼�������޸�!", vbInformation, gstrSysName
        End If
    End Select
    vsList.TextMatrix(Row, col_�޸ı�־) = "Update"
End Sub

Private Sub RecordBuffer(ByVal bytMode As Byte, ByVal vsfVal As VSFlexGrid)
'���ܣ����浱ǰִ����Ŀ��¼����ָ��޸�ǰ��ִ����Ŀ��¼��Ϣ
    Dim i As Integer
    
    If bytMode = 2 Then
        '�ָ�
        For i = 0 To UBound(marrRecord)
            vsfVal.TextMatrix(vsfVal.Row, i) = marrRecord(i)
        Next
    Else
        '����
        If UBound(marrRecord) < 0 Then ReDim Preserve marrRecord(vsfVal.Cols - 1)
        For i = 0 To UBound(marrRecord)
            marrRecord(i) = vsfVal.TextMatrix(vsfVal.Row, i)
        Next
    End If
End Sub




