VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmLisRptGeneral 
   BorderStyle     =   0  'None
   Caption         =   "frmLisStationWrite"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12540
   Icon            =   "frmLisRptGeneral.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTab 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   2280
      Left            =   8775
      ScaleHeight     =   2280
      ScaleWidth      =   3900
      TabIndex        =   33
      Top             =   3195
      Width           =   3900
      Begin XtremeSuiteControls.TabControl TabThis 
         Height          =   2280
         Left            =   75
         TabIndex        =   34
         Top             =   165
         Width           =   3765
         _Version        =   589884
         _ExtentX        =   6641
         _ExtentY        =   4022
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox pic�ٴ����� 
      BorderStyle     =   0  'None
      Height          =   2280
      Left            =   8265
      ScaleHeight     =   2280
      ScaleWidth      =   3900
      TabIndex        =   31
      Top             =   3735
      Width           =   3900
      Begin VB.TextBox txt�ο� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1950
         Left            =   315
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   345
         Width           =   3600
      End
   End
   Begin VB.PictureBox pic��� 
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   8115
      ScaleHeight     =   1185
      ScaleWidth      =   3900
      TabIndex        =   27
      Top             =   1545
      Width           =   3900
      Begin VB.TextBox txt��� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   105
         Width           =   4020
      End
   End
   Begin VB.PictureBox pic��ע 
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   8145
      ScaleHeight     =   1185
      ScaleWidth      =   3900
      TabIndex        =   25
      Top             =   195
      Width           =   3900
      Begin VB.TextBox txt��ע 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   120
         Width           =   4020
      End
   End
   Begin VB.PictureBox picRpt 
      BorderStyle     =   0  'None
      Height          =   4590
      Left            =   60
      ScaleHeight     =   4590
      ScaleWidth      =   8010
      TabIndex        =   12
      Top             =   165
      Width           =   8010
      Begin VB.CheckBox chk��� 
         Appearance      =   0  'Flat
         Caption         =   "���"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4050
         TabIndex        =   30
         Top             =   30
         Width           =   675
      End
      Begin VB.CheckBox chk��ע 
         Appearance      =   0  'Flat
         Caption         =   "��ע"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3380
         TabIndex        =   29
         Top             =   30
         Width           =   675
      End
      Begin VB.CheckBox chkChina 
         Appearance      =   0  'Flat
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   15
         TabIndex        =   17
         Top             =   30
         Width           =   690
      End
      Begin VB.CheckBox chkMB 
         Appearance      =   0  'Flat
         Caption         =   "ø��"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2725
         TabIndex        =   16
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkReferrence 
         Appearance      =   0  'Flat
         Caption         =   "�ο�"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2070
         TabIndex        =   15
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkUnit 
         Appearance      =   0  'Flat
         Caption         =   "��λ"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1415
         TabIndex        =   14
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkSign 
         Appearance      =   0  'Flat
         Caption         =   "��־"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   700
         TabIndex        =   13
         Top             =   30
         Width           =   720
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   4095
         Left            =   0
         TabIndex        =   24
         Top             =   315
         Width           =   7920
         _cx             =   13970
         _cy             =   7223
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483634
         FocusRect       =   2
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
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
         Editable        =   2
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "��ʾ"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   6540
         TabIndex        =   23
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl��ʾ 
         BackColor       =   &H000040C0&
         Height          =   210
         Left            =   6210
         TabIndex        =   22
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ƫ��"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   5760
         TabIndex        =   21
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblƫ�� 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Height          =   210
         Left            =   5430
         TabIndex        =   20
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ƫ��"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   5010
         TabIndex        =   19
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblƫ�� 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Height          =   210
         Left            =   4650
         TabIndex        =   18
         Top             =   45
         Width           =   285
      End
   End
   Begin VB.PictureBox picChart 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   75
      ScaleHeight     =   2565
      ScaleWidth      =   9600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4950
      Width           =   9600
      Begin VB.CommandButton cmdRefersh 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6540
         Picture         =   "frmLisRptGeneral.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   465
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6135
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "3"
         Top             =   90
         Width           =   330
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4545
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "10"
         Top             =   90
         Width           =   375
      End
      Begin VB.OptionButton opt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "���ֵ(&2)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1140
         TabIndex        =   4
         Top             =   75
         Width           =   1125
      End
      Begin VB.OptionButton opt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "������(&1)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   75
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "������(&3)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2250
         TabIndex        =   2
         Top             =   75
         Width           =   1125
      End
      Begin C1Chart2D8.Chart2D chtThis 
         Height          =   1965
         Left            =   60
         TabIndex        =   5
         Top             =   345
         Width           =   8415
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   14843
         _ExtentY        =   3466
         _StockProps     =   0
         ControlProperties=   "frmLisRptGeneral.frx":685E
      End
      Begin VB.Label lbl��Ŀ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����:RBC"
         Height          =   180
         Left            =   7050
         TabIndex        =   10
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ٴ���:"
         Height          =   180
         Left            =   4965
         TabIndex        =   9
         Top             =   90
         Width           =   1170
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������:"
         Height          =   180
         Left            =   3390
         TabIndex        =   8
         Top             =   90
         Width           =   1170
      End
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7545
      Width           =   12540
      _ExtentX        =   22119
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   8115
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLisRptGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ������� = 0: ������Ŀ: Ӣ����: ������: ��λ: CV: �����־: ����ο�:  OD: CUTOFF: COV: С��: ������ĿID: �걾ID: ������� ': ø���ID: ���챨��: ���쾯ʾ '�����Χ: �̶���Ŀ: ��������: ��������: ������Ŀid:  �걾id:
End Enum
Private Const mColCount = 15
Private mstrEndTime As String    '���μ�������
Private mintIdentMode As Integer    '��ʷ�Ƚϲ���ʶ��ʽ
Private mlngҽ��ID As Long
Public mlngMod As Long '����ģ��
Private mrsVsf As ADODB.Recordset

Private Sub chkChina_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("�鿴����", Me.chkChina.value, glngSys, mlngMod)
End Sub

Private Sub chkMB_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("�鿴ø��", Me.chkMB.value, glngSys, mlngMod)
End Sub

Private Sub chkReferrence_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("�鿴�ο�", Me.chkReferrence.value, glngSys, mlngMod)
End Sub

Private Sub chkSign_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("�鿴��־", Me.chkSign.value, glngSys, mlngMod)
End Sub

Private Sub chkUnit_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("�鿴��λ", Me.chkUnit.value, glngSys, mlngMod)
End Sub

Private Sub chk��ע_Click()
    If chk��ע.value = 1 Then
        dkpMan.ShowPane 3
    Else
        dkpMan.FindPane(3).Close
    End If
    Call zlDatabase.SetPara("��ʾ��ע", Me.chk��ע.value, glngSys, mlngMod)
End Sub

Private Sub chk���_Click()
    If chk���.value = 1 Then
        dkpMan.ShowPane 4
    Else
        dkpMan.FindPane(4).Close
    End If
    Call zlDatabase.SetPara("��ʾ���", Me.chk���.value, glngSys, mlngMod)
End Sub

Private Sub cmdRefersh_Click()
    Call vsf_RowColChange
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionCollapsing Or Action = PaneActionCollapsed Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1: Item.Handle = Me.picRpt.Hwnd
'    Case 2: Item.Handle = Me.pic�ο�.hWnd
    Case 3: Item.Handle = Me.pic��ע.Hwnd
    Case 4: Item.Handle = Me.pic���.Hwnd
    Case 5: Item.Handle = Me.picTab.Hwnd
    End Select
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
     Bottom = Me.sbrInfo.Height
End Sub

Private Sub dkpMan_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    'Call RefVsf
End Sub

Private Sub Form_Load()
    '���񻮷�
    '-----------------------------------------------------
    Dim panThis As Pane, pan2 As Pane, pan3 As Pane, pan4 As Pane, Pan5 As Pane
    Set panThis = dkpMan.CreatePane(1, 600, 400, DockTopOf, Nothing)
    panThis.Title = "���鱨��"
    panThis.Options = PaneNoCaption
    
    Set pan3 = dkpMan.CreatePane(3, 200, 400, DockRightOf, panThis)
    pan3.Title = "���鱸ע"
    
    Set pan4 = dkpMan.CreatePane(4, 200, 400, DockBottomOf, pan3)
    pan4.Title = "�����Ϣ"
    
    Set panThis = dkpMan.CreatePane(5, 200, 300, DockBottomOf, Nothing)
    panThis.Title = "��ʷ�Ա�ͼ"
    panThis.Options = PaneNoCaption
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
'    Set pan2 = dkpMan.CreatePane(2, 200, 400, DockRightOf, panThis)
'    pan2.Title = "��ϲο�"
    
    Call initVsf
    Call IntiTab
    
    chkChina.value = Val(zlDatabase.GetPara("�鿴����", glngSys, mlngMod, 1))
    chkSign.value = Val(zlDatabase.GetPara("�鿴��־", glngSys, mlngMod, 1))
    chkUnit.value = Val(zlDatabase.GetPara("�鿴��λ", glngSys, mlngMod, 1))
    chkReferrence.value = Val(zlDatabase.GetPara("�鿴�ο�", glngSys, mlngMod, 1))
    chkMB.value = Val(zlDatabase.GetPara("�鿴ø��", glngSys, mlngMod, 1))
    chk��ע.value = Val(zlDatabase.GetPara("�鿴��ע", glngSys, mlngMod, 1))
    chk���.value = Val(zlDatabase.GetPara("�鿴���", glngSys, mlngMod, 1))
    
    '22539���жϱ�ע������Ƿ񱣴�
    If chk��ע.value = 1 Then
        dkpMan.ShowPane 3
    Else
        dkpMan.FindPane(3).Close
    End If
    If chk���.value = 1 Then
        dkpMan.ShowPane 4
    Else
        dkpMan.FindPane(4).Close
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call zlDatabase.SetPara("�鿴��־", Me.chkSign.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("�鿴��λ", Me.chkUnit.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("�鿴�ο�", Me.chkReferrence.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("�鿴ø��", Me.chkMB.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("�鿴����", Me.chkChina.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("�鿴��ע", Me.chk��ע.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("�鿴���", Me.chk���.value, glngSys, mlngMod)
    mlngҽ��ID = 0
    Set mrsVsf = Nothing
End Sub

Private Sub opt����_Click(Index As Integer)
    Call vsf_RowColChange
End Sub

Private Sub picChart_Resize()
    err = 0: On Error Resume Next
    With Me.chtThis
        
        .Left = 0
        .Width = Me.picChart.ScaleWidth
        .Height = Me.picChart.ScaleHeight - .Top
        
    End With
    
'    chk�ο�.Left = Me.picChart.ScaleWidth - chk�ο�.Width
'    If chk�ο�.Value = 1 Then
'        Me.txt�ο�.Top = Me.chtThis.Top
'        Me.txt�ο�.Left = Me.chtThis.Left + Me.chtThis.Width + 30
'        Me.txt�ο�.Width = Me.picChart.ScaleWidth - Me.chtThis.Width - 30
'        Me.txt�ο�.Height = Me.chtThis.Height
'    End If
End Sub

Private Sub picRpt_Resize()
    On Error Resume Next
    With vsf
        .Top = chkChina.Top + chkChina.Height + 10
        .Left = 10
        .Width = picRpt.ScaleWidth - 20
        .Height = picRpt.Height - .Top - 10
    End With
    Call RefVsf
End Sub

Private Sub picTab_Resize()
    With Me.TabThis
        .Top = 0
        .Left = 0
        .Width = Me.picTab.ScaleWidth
        .Height = Me.picTab.ScaleHeight
    End With
End Sub

Private Sub pic��ע_Resize()
    With Me.txt��ע
    .Left = 0
    .Top = 0
    .Width = Me.pic��ע.ScaleWidth
    .Height = Me.pic��ע.ScaleHeight
    End With
End Sub

Private Sub pic�ٴ�����_Resize()
    With Me.txt�ο�
        .Left = 0
        .Top = 0
        .Width = Me.pic�ٴ�����.ScaleWidth
        .Height = Me.pic�ٴ�����.ScaleHeight
    End With
End Sub

Private Sub pic���_Resize()
    With Me.txt���
        .Left = 0
        .Top = 0
        .Width = Me.pic���.ScaleWidth
        .Height = Me.pic���.ScaleHeight
    End With
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub IntiTab()


    On Error Resume Next

    With Me.TabThis
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearanceExcel
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True

        .PaintManager.ClientFrame = xtpTabFrameSingleLine
'        .PaintManager.Position = xtpTabPositionBottom
        .InsertItem(0, "ͼ������", picChart.Hwnd, conMenu_Tool_Monitor).Tag = "ͼ������"
        .InsertItem(1, "�ٴ�����", pic�ٴ�����.Hwnd, conMenu_View_ToolBar_Text).Tag = "�ٴ�����"
        
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
        
    End With
End Sub

Public Sub zlRefresh(ByVal lngҽ��ID As Long)
    '��ʾ������,��ҽ��ID��ʾ,����һ�����յ������
    
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errHandle
    
    vsf.Rows = 1: vsf.Rows = 2
    Me.txt���.Text = ""
    strSql = "Select ���鱸ע,������,����ʱ��,�����,���ʱ�� From ����걾��¼ where ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngҽ��ID)
    If rsTmp.EOF Then
        Me.txt��ע.Text = ""
        
        With sbrInfo
            .Panels(1).Text = "�����ˣ�"
            .Panels(2).Text = "����ʱ�䣺"
            .Panels(3).Text = "����ˣ�"
            .Panels(4).Text = "���ʱ�䣺"
        End With
        
        Exit Sub
    Else
        Me.txt��ע.Text = Trim("" & rsTmp!���鱸ע)
        
        With sbrInfo
            .Panels(1).Text = "�����ˣ�" & rsTmp!������
            .Panels(2).Text = "����ʱ�䣺" & IIF(IsNull(rsTmp("����ʱ��")), "", Format(rsTmp("����ʱ��"), "yyyy-MM-dd hh:mm"))
            .Panels(3).Text = "����ˣ�" & rsTmp!�����
            .Panels(4).Text = "���ʱ�䣺" & IIF(IsNull(rsTmp("���ʱ��")), "", Format(rsTmp("���ʱ��"), "yyyy-MM-dd hh:mm"))
        End With
    End If
    
    strSql = "Select /*+ rule */" & vbNewLine & _
            "Distinct A.�걾id, A.������Ŀid, A.����, A.�������, A.�̶���Ŀ, A.ID, A.������Ŀ,A.��д as Ӣ����, " & vbNewLine & _
            "         A.Cv, Decode(A.���ν��, '-', '���ԣ�-��', '+', '���ԣ�+��', '*', '*.**', A.���ν��) As ���ν��," & vbNewLine & _
            "         Rownum As ���, A.��־, A.����id, A.�걾���, A.����ʱ��, A.�걾���, A.�걾����ʾ," & vbNewLine & _
            "         A.���鱸ע, A.����, A.�Ա�, A.����, A.�����, A.סԺ��, A.��ǰ����, A.��ҳid, A.�����Χ," & vbNewLine & _
            "         Nvl(G.С��λ��, 2) As С��, A.��������, A.��������, A.��λ," & vbNewLine & _
            "         a.����ο� As �ο�, A.Od," & vbNewLine & _
            "         A.Cutoff, A.Cov, A.ø���id, A.���챨��, A.���쾯ʾ, A.�������,A.����ο�" & vbNewLine & _
            "From (Select A.ID As �걾id, B.������Ŀid, lpad(Decode(D.�������, Null, Nvl(H.����, C.����), D.�������),4,'0') As ����," & vbNewLine & _
            "              Nvl(B.�������, 9999) As �������, Decode(B.������Ŀid, Null, 0, 1) As �̶���Ŀ, B.������Ŀid As ID," & vbNewLine & _
            "              C.������ || Decode(D.��д, Null, '', '(' || D.��д || ')') As ������Ŀ, D.��д,B.ԭʼ���, '' As �ϴν��," & vbNewLine & _
            "              '' As �ϴ�ʱ��, '' As Cv, B.������ As ���ν��, D.���㹫ʽ, D.�������," & vbNewLine & _
            "              Decode(B.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
            "              Nvl(A.����id, -1) As ����id, Nvl(A.�걾���, 0) As �걾���, A.����ʱ��, A.�걾���," & vbNewLine & _
            "              Decode(A.����id, Null," & vbNewLine & _
            "                      To_Char(Trunc(A.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(A.�걾���, 10000), '0000')," & vbNewLine & _
            "                      A.�걾���) As �걾����ʾ, A.���鱸ע, A.����, A.�Ա�, A.����, A.�걾����, A.��������, A.�����," & vbNewLine & _
            "              A.סԺ��, A.���� As ��ǰ����, A.��ҳid, D.�����Χ, D.��������, D.��������, D.��λ, B.Od, B.Cutoff," & vbNewLine & _
            "              B.Sco As Cov, B.ø���id, D.���챨���� As ���챨��, D.���쾯ʾ�� As ���쾯ʾ,B.����ο�" & vbNewLine & _
            "       From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ������Ŀ D, ������ĿĿ¼ H" & vbNewLine & _
            "       Where A.ID = B.����걾id And B.������Ŀid = C.ID And C.ID = D.������Ŀid And B.������Ŀid = H.ID(+) And" & vbNewLine & _
            "             B.��¼���� = A.������ And A.ҽ��ID = [1]"
    strSql = strSql & "       Union All" & vbNewLine & _
            "       Select A.ID As �걾id, B.������Ŀid, lpad(Decode(D.�������, Null, Nvl(H.����, C.����), D.�������),4,'0') As ����," & vbNewLine & _
            "              Nvl(B.�������, 9999) As �������, Decode(B.������Ŀid, Null, 0, 1) As �̶���Ŀ,B.������Ŀid As ID," & vbNewLine & _
            "              C.������ || Decode(D.��д, Null, '', '(' || D.��д || ')') As ������Ŀ,D.��д,B.ԭʼ���, '' As �ϴν��," & vbNewLine & _
            "              '' As �ϴ�ʱ��, '' As Cv, B.������ As ���ν��, D.���㹫ʽ, D.�������," & vbNewLine & _
            "              Decode(B.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
            "              Nvl(A.����id, -1) As ����id, Nvl(A.�걾���, 0) As �걾���, A.����ʱ��, A.�걾���," & vbNewLine & _
            "              Decode(A.����id, Null," & vbNewLine & _
            "                      To_Char(Trunc(A.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(A.�걾���, 10000), '0000')," & vbNewLine & _
            "                      A.�걾���) As �걾����ʾ, A.���鱸ע, A.����, A.�Ա�, A.����, A.�걾����, A.��������, A.�����," & vbNewLine & _
            "              A.סԺ��, A.���� As ��ǰ����, A.��ҳid, D.�����Χ, D.��������, D.��������, D.��λ, B.Od, B.Cutoff," & vbNewLine & _
            "              B.Sco As Cov, B.ø���id, D.���챨���� As ���챨��, D.���쾯ʾ�� As ���쾯ʾ,B.����ο�" & vbNewLine & _
            "       From ����걾��¼ A,����걾��¼ E, ������ͨ��� B, ����������Ŀ C, ������Ŀ D, ����������Ŀ G, ������ĿĿ¼ H" & vbNewLine & _
            "       Where A.ID = B.����걾id And B.������Ŀid = C.ID And C.ID = D.������Ŀid And B.������Ŀid = H.ID(+) And" & vbNewLine & _
            "             B.��¼���� = A.������ And E.ID=A.�ϲ�id  And E.ҽ��ID= [1]) A, ����������Ŀ G" & vbNewLine & _
            "Where A.����id = G.����id(+) And A.ID = G.��Ŀid(+)" & vbNewLine & _
            "Order By A.����, A.�������"
    
    Set mrsVsf = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngҽ��ID)
    
    Call RefVsf
    '��ʾ�����Ϣ
    Dim strTmp As String
    strSql = "Select distinct b.ҽ��id, b.��Ŀ, b.����, b.����" & vbNewLine & _
                "From ����걾��¼ a, ����ҽ������ b" & vbNewLine & _
                "Where a.ҽ��id = b.ҽ��id and a.ҽ��ID = [1] " & vbNewLine & _
                "Order By ҽ��id, ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngҽ��ID)
    Do Until rsTmp.EOF
        strTmp = strTmp & Trim("" & rsTmp("��Ŀ")) & ":" & Replace(Trim("" & rsTmp("����")), vbCrLf, vbCrLf & "    ") & vbCrLf
        rsTmp.MoveNext
    Loop
    Me.txt���.Text = strTmp
    
    Call RefChartData(lngҽ��ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefVsf()
    Dim lngRow As Long, lngCol As Long
    Dim bln���� As Boolean
    Call initVsf
    If mrsVsf Is Nothing Then Exit Sub
    
    lngRow = vsf.FixedRows
    If mrsVsf.RecordCount > 0 Then mrsVsf.MoveFirst
    Do Until mrsVsf.EOF
        With vsf
            Dim lngAdd As Long
            lngAdd = 1
            'If .ScrollBars >= flexScrollBarVertical Then lngAdd = 2
            If (Split(Format(.ClientHeight / .RowHeightMin, "0.0000"), ".")(0)) >= lngRow + 2 And bln���� = False Then
                lngCol = 0
            Else
                If lngRow > 5 Then '����5��
                    If bln���� = False Then
                        bln���� = True
                        Call Add_Column(lngCol, lngRow)
                    End If
                End If
            End If
            If (lngCol = 0 And lngRow >= .Rows) Or (lngCol > 0 And lngRow >= .Rows - 1) Then
                Call Add_Column(lngCol, lngRow)
            End If

            .TextMatrix(lngRow, mCol.������� + lngCol * mColCount) = mrsVsf.Bookmark  'Trim("" & mrsVsf!�������)
            .TextMatrix(lngRow, mCol.������Ŀ + lngCol * mColCount) = Trim("" & mrsVsf!������Ŀ)
            .TextMatrix(lngRow, mCol.Ӣ���� + lngCol * mColCount) = Trim("" & mrsVsf!Ӣ����)
            .TextMatrix(lngRow, mCol.������ + lngCol * mColCount) = Trim("" & mrsVsf!���ν��)
            .TextMatrix(lngRow, mCol.��λ + lngCol * mColCount) = Trim("" & mrsVsf!��λ)
            .TextMatrix(lngRow, mCol.CV + lngCol * mColCount) = Trim("" & mrsVsf!CV)
            .TextMatrix(lngRow, mCol.�����־ + lngCol * mColCount) = Trim("" & mrsVsf!��־)
            .TextMatrix(lngRow, mCol.����ο� + lngCol * mColCount) = IIF(Trim("" & mrsVsf!�ο�) = "", Trim("" & mrsVsf!����ο�), Trim("" & mrsVsf!�ο�))
            .TextMatrix(lngRow, mCol.OD + lngCol * mColCount) = Trim("" & mrsVsf!OD)
            .TextMatrix(lngRow, mCol.CUTOFF + lngCol * mColCount) = Trim("" & mrsVsf!CUTOFF)
            .TextMatrix(lngRow, mCol.COV + lngCol * mColCount) = Trim("" & mrsVsf!COV)
            .TextMatrix(lngRow, mCol.С�� + lngCol * mColCount) = Trim("" & mrsVsf!С��)
            .TextMatrix(lngRow, mCol.������ĿID + lngCol * mColCount) = Trim("" & mrsVsf!ID)
            .TextMatrix(lngRow, mCol.�걾ID + lngCol * mColCount) = Trim("" & mrsVsf!�걾ID)
            .TextMatrix(lngRow, mCol.������� + lngCol * mColCount) = Trim("" & mrsVsf!�������)
            
            If lngCol = 0 Then
                .Rows = .Rows + 1
            End If
            lngRow = lngRow + 1
        End With
        mrsVsf.MoveNext
    Loop
    Call Check_ColWidth
    
    vsf.Rows = vsf.Rows - 1
End Sub
Private Sub RefChartData(ByVal lngҽ��ID As Long)
    '������ʾ �����ٴ���
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mlngҽ��ID = lngҽ��ID Then Exit Sub
    mlngҽ��ID = lngҽ��ID
    strSql = "Select /* +rule */" & vbNewLine & _
        " Nvl(L.����ʱ��, Sysdate) As ����ʱ��, Nvl(Max(��������), 0) As ����" & vbNewLine & _
        "From ������Ŀѡ�� O, ���鱨����Ŀ X, ������ͨ��� R, ����걾��¼ L" & vbNewLine & _
        "Where O.������Ŀid(+) = X.������Ŀid And X.������Ŀid = R.������Ŀid And R.����걾id = L.ID And L.ҽ��ID = [1]" & vbNewLine & _
        "Group By Nvl(L.����ʱ��, Sysdate)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngҽ��ID)
    Me.txt����.Text = 30
    mstrEndTime = Format(Now(), "yyyy-MM-dd hh:mm:ss")
    Do Until rsTmp.EOF
         Me.txt����.Text = rsTmp!����
         mstrEndTime = Format(rsTmp!����ʱ��, "yyyy-MM-dd hh:mm:ss")
        rsTmp.MoveNext
    Loop
    If Val(Me.txt����.Text) <= 0 Then Me.txt����.Text = 30
    If Val(Me.txt����.Text) <= 0 Then Me.txt���� = 3
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Add_Column(ByRef lngCol As Long, ByRef lngRow As Long)
    '��ӷ����������
    With vsf
        lngCol = lngCol + 1
        lngRow = vsf.FixedRows
        .Cols = .Cols + mColCount
        .TextMatrix(0, mCol.������� + lngCol * mColCount) = "": .ColWidth(mCol.������� + lngCol * mColCount) = 300: .ColAlignment(mCol.������� + lngCol * mColCount) = flexAlignRightCenter
        .TextMatrix(0, mCol.������Ŀ + lngCol * mColCount) = "������Ŀ": .ColWidth(mCol.������Ŀ + lngCol * mColCount) = 2100: .ColAlignment(mCol.������Ŀ + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.Ӣ���� + lngCol * mColCount) = "������Ŀ": .ColWidth(mCol.Ӣ���� + lngCol * mColCount) = 1000: .ColAlignment(mCol.Ӣ���� + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.������ + lngCol * mColCount) = "������": .ColWidth(mCol.������ + lngCol * mColCount) = 1200: .ColAlignment(mCol.������ + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.��λ + lngCol * mColCount) = "��λ": .ColWidth(mCol.��λ + lngCol * mColCount) = 1000: .ColAlignment(mCol.��λ + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.CV + lngCol * mColCount) = "CV": .ColWidth(mCol.CV + lngCol * mColCount) = 0: .ColAlignment(mCol.CV + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.CV + lngCol * mColCount) = True
        .TextMatrix(0, mCol.�����־ + lngCol * mColCount) = "��־": .ColWidth(mCol.�����־ + lngCol * mColCount) = 450: .ColAlignment(mCol.�����־ + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.����ο� + lngCol * mColCount) = "�ο�": .ColWidth(mCol.����ο� + lngCol * mColCount) = 1300: .ColAlignment(mCol.����ο� + lngCol * mColCount) = flexAlignLeftCenter
        
        .TextMatrix(0, mCol.OD + lngCol * mColCount) = "OD": .ColWidth(mCol.OD + lngCol * mColCount) = 700: .ColAlignment(mCol.OD + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.CUTOFF + lngCol * mColCount) = "CUTOFF": .ColWidth(mCol.CUTOFF + lngCol * mColCount) = 700: .ColAlignment(mCol.CUTOFF + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.COV + lngCol * mColCount) = "COV": .ColWidth(mCol.COV + lngCol * mColCount) = 700: .ColAlignment(mCol.COV + lngCol * mColCount) = flexAlignLeftCenter
        
        .TextMatrix(0, mCol.С�� + lngCol * mColCount) = "С��": .ColWidth(mCol.С�� + lngCol * mColCount) = 0: .ColAlignment(mCol.С�� + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.С�� + lngCol * mColCount) = True
        .TextMatrix(0, mCol.������ĿID + lngCol * mColCount) = "������Ŀid": .ColWidth(mCol.������ĿID + lngCol * mColCount) = 0: .ColAlignment(mCol.������ĿID + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.������ĿID + lngCol * mColCount) = True
        .TextMatrix(0, mCol.�걾ID + lngCol * mColCount) = "�걾ID": .ColWidth(mCol.�걾ID + lngCol * mColCount) = 0: .ColAlignment(mCol.�걾ID + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.�걾ID + lngCol * mColCount) = True
        .TextMatrix(0, mCol.������� + lngCol * mColCount) = "�������": .ColWidth(mCol.������� + lngCol * mColCount) = 0: .ColAlignment(mCol.������� + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.������� + lngCol * mColCount) = True
        
        
    End With
End Sub
Private Sub initVsf()
    '��ʼ�����
    With vsf
        .BackColor = &H80000005
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .BackColorFixed = &HFDD6C6
        .GridLinesFixed = flexGridFlat
        .RowHeightMin = 300
        .Editable = flexEDNone
        
        .Rows = 2: .FixedRows = 1
        .Cols = mColCount: .FixedCols = 0
        
        .TextMatrix(0, mCol.�������) = "": .ColWidth(mCol.�������) = 300: .ColAlignment(mCol.�������) = flexAlignRightCenter
        .TextMatrix(0, mCol.������Ŀ) = "������Ŀ": .ColWidth(mCol.������Ŀ) = 2100: .ColAlignment(mCol.������Ŀ) = flexAlignLeftCenter
        .TextMatrix(0, mCol.Ӣ����) = "������Ŀ": .ColWidth(mCol.������Ŀ) = 1000: .ColAlignment(mCol.Ӣ����) = flexAlignLeftCenter
        
        .TextMatrix(0, mCol.������) = "������": .ColWidth(mCol.������) = 1200: .ColAlignment(mCol.������) = flexAlignLeftCenter
        .TextMatrix(0, mCol.��λ) = "��λ": .ColWidth(mCol.��λ) = 1000: .ColAlignment(mCol.��λ) = flexAlignLeftCenter
        .TextMatrix(0, mCol.CV) = "CV": .ColWidth(mCol.CV) = 0: .ColAlignment(mCol.CV) = flexAlignLeftCenter
        .ColHidden(mCol.CV) = True
        .TextMatrix(0, mCol.�����־) = "��־": .ColWidth(mCol.�����־) = 450: .ColAlignment(mCol.�����־) = flexAlignLeftCenter
        .TextMatrix(0, mCol.����ο�) = "�ο�": .ColWidth(mCol.����ο�) = 1300: .ColAlignment(mCol.����ο�) = flexAlignLeftCenter


        .TextMatrix(0, mCol.OD) = "OD": .ColWidth(mCol.OD) = 700: .ColAlignment(mCol.OD) = flexAlignLeftCenter
        .TextMatrix(0, mCol.CUTOFF) = "CUTOFF": .ColWidth(mCol.CUTOFF) = 700: .ColAlignment(mCol.CUTOFF) = flexAlignLeftCenter
        .TextMatrix(0, mCol.COV) = "COV": .ColWidth(mCol.COV) = 700: .ColAlignment(mCol.COV) = flexAlignLeftCenter
        .TextMatrix(0, mCol.С��) = "С��": .ColWidth(mCol.С��) = 0: .ColAlignment(mCol.С��) = flexAlignLeftCenter
        .ColHidden(mCol.С��) = True
        .TextMatrix(0, mCol.������ĿID) = "������ĿID": .ColWidth(mCol.������ĿID) = 0: .ColAlignment(mCol.������ĿID) = flexAlignLeftCenter
        .ColHidden(mCol.������ĿID) = True
        .TextMatrix(0, mCol.�걾ID) = "�걾ID": .ColWidth(mCol.�걾ID) = 0: .ColAlignment(mCol.�걾ID) = flexAlignLeftCenter
        .ColHidden(mCol.�걾ID) = True
        .TextMatrix(0, mCol.�������) = "�������": .ColWidth(mCol.�������) = 0: .ColAlignment(mCol.�������) = flexAlignLeftCenter
        .ColHidden(mCol.�������) = True
        
        Call Check_ColWidth
    End With
End Sub

Private Sub Check_ColWidth()
    '���ݿؼ�״̬�������п�
    
    Dim lngCol As Long, lngLoop As Long, lngRow As Long
    Dim lngColor As Long, lngForeColor As Long, str��־ As String
    With vsf
        lngCol = (.Cols / mColCount)
        For lngLoop = 0 To lngCol - 1
            '����е���ɫ
            .Cell(flexcpBackColor, 1, mCol.������� + lngLoop * mColCount, vsf.Rows - 1, mCol.������� + lngLoop * mColCount) = vsf.BackColorFixed
            
            '�п�
            .ColWidth(mCol.������Ŀ + lngLoop * mColCount) = IIF(chkChina.value = 0, 0, 2100)
            .ColWidth(mCol.Ӣ���� + lngLoop * mColCount) = IIF(chkChina.value = 0, 1000, 0)
            .ColWidth(mCol.�����־ + lngLoop * mColCount) = IIF(chkSign.value = 0, 0, 450)
            .ColWidth(mCol.��λ + lngLoop * mColCount) = IIF(chkUnit.value = 0, 0, 1000)
            .ColWidth(mCol.����ο� + lngLoop * mColCount) = IIF(chkReferrence.value = 0, 0, 1300)
            
            .ColWidth(mCol.OD + lngLoop * mColCount) = IIF(chkMB.value = 0, 0, 700)
            .ColWidth(mCol.CUTOFF + lngLoop * mColCount) = IIF(chkMB.value = 0, 0, 700)
            .ColWidth(mCol.COV + lngLoop * mColCount) = IIF(chkMB.value = 0, 0, 700)
            
            For lngRow = .FixedRows To .Rows - 1
                '��Ԫ���ʽ
                If IsNumeric("-" & .TextMatrix(lngRow, mCol.������ + lngLoop * mColCount)) Then
                    .TextMatrix(lngRow, mCol.������ + lngLoop * mColCount) = Format(.TextMatrix(lngRow, mCol.������ + lngLoop * mColCount), _
                        IIF(Val(.TextMatrix(lngRow, mCol.С�� + lngLoop * mColCount)) = 0, "#0", "0." & String(Val(.TextMatrix(lngRow, mCol.С�� + lngLoop * mColCount)), "0")))
                End If
                '��ɫ
                lngColor = .BackColor
                lngForeColor = .ForeColor
                str��־ = Trim(.TextMatrix(lngRow, mCol.�����־ + lngLoop * mColCount))
                If InStr("��", str��־) > 0 And str��־ <> "" Then     '2
                    lngColor = lblƫ��.BackColor
                    lngForeColor = lblƫ��.ForeColor
                ElseIf InStr("��,�쳣", str��־) > 0 And str��־ <> "" Then '3,�쳣
                    lngColor = lblƫ��.BackColor
                    lngForeColor = lblƫ��.ForeColor
                ElseIf InStr("����,����", str��־) > 0 And str��־ <> "" Then '5,6
                    lngColor = lbl��ʾ.BackColor
                    lngForeColor = lbl��ʾ.ForeColor
                End If
                .Cell(flexcpBackColor, lngRow, mCol.������ + lngLoop * mColCount, lngRow, mCol.������ + lngLoop * mColCount) = lngColor
                .Cell(flexcpForeColor, lngRow, mCol.������ + lngLoop * mColCount, lngRow, mCol.������ + lngLoop * mColCount) = lngForeColor
            Next
        Next
    End With
End Sub

Private Function get_Column(ByVal lngCol As Long) As Long
    '�õ�ָ�����ǵڼ�������
    Dim strTmp As String
    strTmp = CStr(Format(lngCol / mColCount, "0.00000"))
    If InStr(strTmp, ".") > 0 Then
        get_Column = Val(Mid(strTmp, 1, InStr(strTmp, ".")))
    Else
        get_Column = Val(strTmp)
    End If
End Function

Private Sub vsf_RowColChange()
    Dim lng��Ŀid As Long, str��Ŀ As String, str������� As String, lng�걾id As Long
    Dim lngCol As Long
    
    lngCol = get_Column(vsf.Col)
    
    lng��Ŀid = Val(vsf.TextMatrix(vsf.Row, mCol.������ĿID + lngCol * mColCount))
    lng�걾id = Val(vsf.TextMatrix(vsf.Row, mCol.�걾ID + lngCol * mColCount))
    If chkChina.value Then
        str��Ŀ = vsf.TextMatrix(vsf.Row, mCol.������Ŀ + lngCol * mColCount)
    Else
        str��Ŀ = vsf.TextMatrix(vsf.Row, mCol.Ӣ���� + lngCol * mColCount)
    End If
    Me.lbl��Ŀ.Caption = "��Ŀ��" & str��Ŀ
    Me.lbl��Ŀ.Left = Me.cmdRefersh.Left + Me.cmdRefersh.Width + 45
    
    str������� = vsf.TextMatrix(vsf.Row, mCol.������� + lngCol * mColCount)
    If lng��Ŀid <> 0 Then
        Call RefChart(lng�걾id, lng��Ŀid, str�������)
        
    End If
End Sub


Private Sub RefChart(ByVal lng�걾id As Long, ByVal lng��Ŀid As Long, ByVal str������� As String)
    '������ͼ
    Dim aryX() As Variant, aryY() As Variant
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lngDates As Long, lngCount As Long, intLoop As Integer, dblAvg As Double, lng���� As Long
    Dim strMaxValue As String, strMinValue As String
    Dim dblCurCV As Double, dbl���ν�� As Double, dbl���챨���� As Double
    On Error GoTo errHandle
    '��������������Ϊ0�����ͼ����ʾ
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    
    
    '�ٴ�����
    txt�ο�.Text = ""
    strSql = "Select �ٴ����� From ������Ŀ Where ������Ŀid=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng��Ŀid)
    Do Until rsTmp.EOF
        txt�ο�.Text = Trim("" & rsTmp.Fields("�ٴ�����"))
        rsTmp.MoveNext
    Loop
    
    
    If str������� = "2" Or str������� = "3" Then
       Me.chtThis.IsBatched = False
       Exit Sub
    End If
    
    '����ͼ�εĻ�����̬
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot  '����
        .Styles(oc2dTypePlot).Symbol.Shape = oc2dShapeBox
        With .Data
            .Layout = oc2dDataArray
            .NumSeries = 1
            .NumPoints(1) = 4
        End With
    End With
    With Me.chtThis.ChartArea
        .Axes("X").MajorGrid.Spacing.IsDefault = True
        .Axes("Y").MajorGrid.Spacing.IsDefault = True
        .Axes("X").AnnotationMethod = oc2dAnnotateValueLabels   '��������ʾֵ��ʾ
    End With
    
    If Me.opt����(0).value = True Then
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "������"
    ElseIf Me.opt����(1).value = True Then
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "���ֵ"
    Else
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "������"
    End If
    
    '������
    lngDates = Val(Me.txt����.Text)
    lng���� = Val(Me.txt����.Text)
    
    If Me.opt����(2).value = True Then
        strSql = "Select ����, ������, Ӣ����, ������Ŀid, ������, decode(����,null,������,����) as ���� " & vbNewLine & _
                    "From (Select Decode(E.�������, Null, D.����, E.�������) As ����, D.������, D.Ӣ����, B.������Ŀid, B.������, H.����" & vbNewLine & _
                    "       From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ����������Ŀ D, ������Ŀ E, ���鱨����Ŀ F, ������ĿĿ¼ G," & vbNewLine & _
                    "            (Select ������Ŀid, ���� As ���� From ������Ŀ���� Where ���� = 9 And ���� = 1) H" & vbNewLine & _
                    "       Where A.ID = B.����걾id And B.������Ŀid = C.��Ŀid And B.������Ŀid = D.ID And Nvl(C.��������Ŀ, 0) = -1 And A.ҽ��ID = [1] And" & vbNewLine & _
                    "             B.������Ŀid = E.������Ŀid And B.������Ŀid = F.������Ŀid And F.������Ŀid = G.ID And Nvl(G.�����Ŀ, 0) = 0 And" & vbNewLine & _
                    "             G.ID = H.������Ŀid(+)" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(E.�������, Null, D.����, E.�������) As ����, D.������, D.Ӣ����, B.������Ŀid, B.������, H.����" & vbNewLine & _
                    "       From ����걾��¼ A, ����걾��¼ I,������ͨ��� B, ����������Ŀ C, ����������Ŀ D, ������Ŀ E, ���鱨����Ŀ F, ������ĿĿ¼ G," & vbNewLine & _
                    "            (Select ������Ŀid, ���� As ���� From ������Ŀ���� Where ���� = 9 And ���� = 1) H" & vbNewLine & _
                    "       Where A.ID = B.����걾id And B.������Ŀid = C.��Ŀid And B.������Ŀid = D.ID And Nvl(C.��������Ŀ, 0) = -1 And A.�ϲ�id = I.ID And I.ҽ��ID= [1] And" & vbNewLine & _
                    "             B.������Ŀid = E.������Ŀid And B.������Ŀid = F.������Ŀid And F.������Ŀid = G.ID And Nvl(G.�����Ŀ, 0) = 0 And" & vbNewLine & _
                    "             G.ID = H.������Ŀid(+))" & vbNewLine & _
                    "Order By ����"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngҽ��ID)
        If rsTmp.RecordCount > 0 Then
            ReDim aryX(rsTmp.RecordCount - 1)
            ReDim aryY(rsTmp.RecordCount - 1, 0)
        End If
    Else
        strSql = "Select I.ID, I.���� As ������, V.��д As Ӣ����, I.���㵥λ As ��λ, L.����, L.����ʱ��, L.������, V.���챨����,V.������� " & vbNewLine & _
                "From (Select L.������Ŀid, L.����, L.����ʱ��, L.������ " & vbNewLine & _
                "       From (Select M.����id As ����id, M.����, M.�Ա�, L.ID As ����, L.����ʱ��, R.������Ŀid, R.������,L.�걾���� " & vbNewLine & _
                "              From ����걾��¼ L, ������ͨ��� R, ����ҽ����¼ M, (select ����id,����,�Ա� from ����걾��¼ where ҽ��id = [1]) N " & vbNewLine & _
                "              Where M.ID = L.ҽ��id And L.ID = R.����걾id And  " & vbNewLine & _
                "                    L.����ʱ�� Between [2]  And" & vbNewLine & _
                "                    [3] and " & IIF(mintIdentMode = 0, "L.����id = N.����id", "L.���� = N.���� And L.�Ա� = N.�Ա�") & ") L," & vbNewLine & _
                "            (Select M.����id As ����id, M.����, M.�Ա�, L.����ʱ��, R.������Ŀid,L.�걾���� " & vbNewLine & _
                "              From ����ҽ����¼ M, ����걾��¼ L, ������ͨ��� R" & vbNewLine & _
                "              Where M.ID = L.ҽ��id And L.ID = R.����걾id And L.ҽ��id = [1]) C" & vbNewLine & _
                "       Where " & IIF(mintIdentMode = 0, "L.����id = C.����id", "L.���� = C.���� And L.�Ա� = C.�Ա�") & " And L.������Ŀid+0 = C.������Ŀid " & _
                "       And L.�걾���� = C.�걾���� ) L, ������Ŀ V, ���鱨����Ŀ R, ������ĿĿ¼ I" & vbNewLine & _
                "Where L.������Ŀid=[4] and L.������Ŀid = V.������Ŀid And L.������Ŀid = R.������Ŀid And R.������Ŀid = I.ID And I.�����Ŀ <> 1" & vbNewLine & _
                "Order By I.����, L.����ʱ�� desc"
                
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngҽ��ID, CDate(Format(mstrEndTime, "yyyy-MM-dd 00:00:00")) - lngDates, _
                                           CDate(Format(mstrEndTime, "yyyy-MM-dd 23:59:59")), lng��Ŀid)
        If rsTmp.RecordCount > 0 Then
            ReDim aryX(rsTmp.RecordCount - 1)
            ReDim aryY(rsTmp.RecordCount - 1, 0)
            If lng���� >= rsTmp.RecordCount Then
                ReDim aryX(rsTmp.RecordCount - 1)
                ReDim aryY(rsTmp.RecordCount - 1, 0)
            Else
                ReDim aryX(lng���� - 1)
                ReDim aryY(lng���� - 1, 0)
            End If
        End If
    End If
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    
    '�������
    If rsTmp.RecordCount > 0 Then
        For lngCount = LBound(aryX) To UBound(aryX)
            
            aryX(lngCount) = lngCount
            If Me.opt����(0).value = True Then  '������
                If lng�걾id = rsTmp!���� Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, "���ν��"
                    dbl���ν�� = Val("" & rsTmp!������)
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                End If
                
                If Val(Trim("" & rsTmp!������)) = 0 Or dbl���ν�� = 0 Then
                    aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                Else
                    dblCurCV = Format((Val(Trim("" & rsTmp!������)) - dbl���ν��) / dbl���ν�� * 100, "0.00")
                    aryY(lngCount, 0) = dblCurCV
                End If
                'Debug.Print "�����" & Val(Trim("" & rsTmp!������)) & ", ������:" & dblCurCV
                dbl���챨���� = Val("" & rsTmp!���챨����)
                
            ElseIf Me.opt����(1).value = True Then '���ֵ
                If lng�걾id = rsTmp!���� Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, "���ν��"
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                End If
                If Val(Trim("" & rsTmp!������)) = 0 Then
                    aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                Else
                    aryY(lngCount, 0) = Val(Trim("" & rsTmp!������))
                End If
            Else                                '������
                Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, Trim("" & rsTmp("����"))
                aryY(lngCount, 0) = Val("" & rsTmp("������"))
                
            End If
            
            rsTmp.MoveNext
            
            If Val(strMaxValue) < Abs(Val(aryY(lngCount, 0))) Then
                strMaxValue = Abs(Val(aryY(lngCount, 0)))
            End If
            If Val(strMinValue) > Abs(Val(aryY(lngCount, 0))) Then
                strMinValue = Abs(Val(aryY(lngCount, 0)))
            End If
            
        Next
        
        '���ˢ���ڲ�����
        Me.chtThis.IsBatched = True
        Me.chtThis.ChartGroups(1).Data.NumPoints(1) = UBound(aryX) + 1
        Call Me.chtThis.ChartGroups(1).Data.CopyXVectorIn(1, aryX)
        Call Me.chtThis.ChartGroups(1).Data.CopyYArrayIn(aryY)
        
        If opt����(0).value = True Then
            Me.chtThis.ChartArea.Axes("Y").Origin = 0
            Me.chtThis.ChartArea.Axes("Y").Min = -1 * Val(strMaxValue)
            Me.chtThis.ChartArea.Axes("Y").Max = Val(strMaxValue)
        ElseIf opt����(1).value = True Then
            On Error Resume Next
            For intLoop = 0 To UBound(aryY, 1) - 1
                dblAvg = dblAvg + Val(aryY(intLoop, 0))
            Next
            If dblAvg <> 0 Then
                dblAvg = dblAvg / UBound(aryY, 1)
                Me.chtThis.ChartArea.Axes("Y").Origin = dblAvg
                If (dblAvg - Val(strMinValue)) < (Val(strMaxValue) - dblAvg) Then
                    Me.chtThis.ChartArea.Axes("Y").Min = Val(dblAvg - (Val(strMaxValue) - dblAvg))
                    Me.chtThis.ChartArea.Axes("Y").Max = Val(dblAvg + (Val(strMaxValue) - dblAvg))
                Else
                    Me.chtThis.ChartArea.Axes("Y").Min = Val(dblAvg - (dblAvg - Val(strMinValue)))
                    Me.chtThis.ChartArea.Axes("Y").Max = Val(dblAvg + (dblAvg - Val(strMinValue)))
                End If
            End If
        Else
            Me.chtThis.ChartArea.Axes("Y").Origin = 0
            Me.chtThis.ChartArea.Axes("Y").Min = 0
            Me.chtThis.ChartArea.Axes("Y").Max = Val(strMaxValue)
        End If
    End If
    Me.chtThis.IsBatched = False
    
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


