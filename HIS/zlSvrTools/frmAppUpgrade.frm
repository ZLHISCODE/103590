VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmAppUpgrade 
   BackColor       =   &H80000005&
   Caption         =   "ϵͳ��Ǩ����"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmAppUpgrade.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   10155
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   0
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   9735
      TabIndex        =   5
      Top             =   600
      Width           =   9732
      Begin VB.CommandButton cmdkillProcess 
         Caption         =   "�жϿͻ������ӵĽ��̶���(&P)"
         Height          =   350
         Left            =   6315
         TabIndex        =   32
         Top             =   1320
         Width           =   2790
      End
      Begin VB.CommandButton cmdSelALl 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   120
         TabIndex        =   30
         Top             =   5160
         Width           =   1100
      End
      Begin VB.CommandButton cmdNotSel 
         Caption         =   "ȫ��(&R)"
         Height          =   350
         Left            =   1200
         TabIndex        =   29
         Top             =   5160
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar pgbGetUpInfo 
         Height          =   180
         Left            =   4200
         TabIndex        =   28
         Top             =   2760
         Visible         =   0   'False
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Frame fraUpMode 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   960
         TabIndex        =   24
         Top             =   2700
         Width           =   2175
         Begin VB.OptionButton optUpMode 
            BackColor       =   &H80000005&
            Caption         =   "��ǰ��Ǩ"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   27
            Top             =   60
            Width           =   1215
         End
         Begin VB.OptionButton optUpMode 
            BackColor       =   &H80000005&
            Caption         =   "������Ǩ"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   26
            Top             =   60
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdExec 
         Caption         =   "ִ��(&E)��"
         Height          =   350
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5160
         Width           =   1100
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   2
         Left            =   0
         TabIndex        =   22
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   3
         Left            =   1020
         TabIndex        =   12
         Top             =   2115
         Width           =   5940
      End
      Begin VB.CommandButton cmdFunc 
         Appearance      =   0  'Flat
         Caption         =   "���ò��жϿͻ���"
         Height          =   350
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   1660
      End
      Begin VB.CommandButton cmdFunc 
         Appearance      =   0  'Flat
         Caption         =   "�����û��ʺ�"
         Height          =   350
         Index           =   1
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   1300
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   1
         Left            =   900
         TabIndex        =   8
         Top             =   1155
         Width           =   5940
      End
      Begin VB.CommandButton cmdFunc 
         Appearance      =   0  'Flat
         Caption         =   "���ô�����"
         Height          =   350
         Index           =   3
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   1100
      End
      Begin VB.CommandButton cmdFunc 
         Appearance      =   0  'Flat
         Caption         =   "���ú�̨��ҵ"
         Height          =   350
         Index           =   2
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   1300
      End
      Begin VSFlex8Ctl.VSFlexGrid vsSysSel 
         Height          =   1812
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   9372
         _cx             =   16531
         _cy             =   3196
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   8
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAppUpgrade.frx":04F9
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
      Begin MSComctlLib.ProgressBar pgbState 
         Height          =   180
         Left            =   960
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComDlg.CommonDialog cdgPub 
         Left            =   3600
         Top             =   1980
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblConfigureFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ǩ������Ҫ�����ű�����Ҫ����Ǩ��Ŀ��ϵͳ�İ�װ�����ļ����Ի�ȡ��Ӧ�Ĺ����߽ű���"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3000
         Width           =   7575
      End
      Begin VB.Label lblUpMode 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��Ǩģʽ"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڼ��"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblExplain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAppUpgrade.frx":05E1
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   10080
      End
      Begin VB.Label lblMainPath 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "ϵͳ��װĿ¼��C:\Appsoft"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   2370
         Width           =   2160
      End
      Begin VB.Label lblSel 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���ġ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   2460
         TabIndex        =   14
         Top             =   2370
         Width           =   540
      End
      Begin VB.Label lblUpgrade 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��Ǩִ��"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblManage 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��Ǩ׼��"
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   720
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5052
      Index           =   1
      Left            =   3660
      ScaleHeight     =   5055
      ScaleWidth      =   9615
      TabIndex        =   1
      Top             =   1515
      Width           =   9612
      Begin VB.ComboBox cboSys 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   4560
      End
      Begin VSFlex8Ctl.VSFlexGrid vsUpLog 
         Height          =   3708
         Left            =   120
         TabIndex        =   2
         Top             =   828
         Width           =   9372
         _cx             =   16531
         _cy             =   6540
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
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   2
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
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAppUpgrade.frx":0686
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblSys 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ��ϵͳ"
         Height          =   180
         Left            =   165
         TabIndex        =   4
         Top             =   360
         Width           =   720
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   6240
      Left            =   0
      TabIndex        =   20
      Top             =   480
      Width           =   9780
      _Version        =   589884
      _ExtentX        =   17251
      _ExtentY        =   11007
      _StockProps     =   64
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳ��Ǩ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   0
      Top             =   105
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   0
      Picture         =   "frmAppUpgrade.frx":0761
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmAppUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum SysSelCol
    Col_Sel = 0
    Col_��� = 1
    Col_���� = 2
    Col_�����ļ� = 3
    Col_��ǰ�汾 = 4
    Col_Ŀ��汾 = 5
    Col_����� = 6
End Enum

Private Enum SysUpCol
    Col_��� = 0
    Col_��Ǩʱ�� = 1
    Col_ԭʼ�汾 = 2
    Col_Ԥ��Ŀ�� = 3
    Col_����汾 = 4
    Col_��Ǩ��� = 5
    Col_��ǰִ�� = 6
End Enum

Private Enum FuncCtrlType
    FCT_�ͻ��� = 0
    FCT_�û��˻� = 1
    FCT_��̨��ҵ = 2
    FCT_������ = 3
End Enum

Private mrsSysInfo As ADODB.Recordset
Private mrsSysUpFiles As ADODB.Recordset
Private mrsMainPath As ADODB.Recordset

Private mstrSysJobs As String  '�ֹ����õ�ϵͳ����
Private mblnLoadSysFiles As Boolean '�Ƿ��Ѿ�����ZLSysFiles�е������ļ�
Private mblnLastUpInfo As Boolean '�Ƿ��ȡ�ϴ���Ǩ��ʷ
'Private mstrMaxUpVer As String '����ϵͳ��������ʱ�����ǰ汾��Ӧʱ�����汾��
Private mobjOprateLog As TextStream
'===========================================================================
'==�����ӿ�
'===========================================================================
Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
End Sub

'===========================================================================
'==�¼�
'===========================================================================
Private Sub cboSys_Click()
    Call LoadData(1)
End Sub

Private Sub cmdExec_Click()
    Dim objfrmUpSys As frmAppUpgradeNew
    Dim strRunModule As String
    
    'ϵͳ��Ϣ��¼��
    Call RecToLog(mrsSysInfo, "ϵͳ���", "ԭʼϵͳϵͳ��¼��")
    If VilidateUpgrade Then
        'ҵ��߷��ڼ��
        If optUpMode(1).value Then
            If Not CheckRushHours("0102", "��ǰ��Ǩ") Then
                Exit Sub
            End If
        End If
        If MsgBox("���鵱ǰ��ѡ������ϵͳ�Ƿ����Ҫ���Ƿ�ȷ��������", vbInformation + vbYesNo, gstrSysName) = vbYes Then
            If Not CheckCBOPars Then Exit Sub
            Set objfrmUpSys = New frmAppUpgradeNew '�������ģ�����
            Call objfrmUpSys.ShowMe(frmMDIMain, mrsSysInfo, mrsSysUpFiles, optUpMode(1).value, strRunModule)
            If strRunModule <> "" Then
                Unload Me
                Call frmMDIMain.RunByModule(strRunModule)
                Exit Sub
            End If
            vsSysSel.Tag = ""
            Call LoadSystems
            Call LoadData(0)
            Call VilidateUpgrade
        End If
    Else
        MsgBox "��ѡ����Ҫ������Ӧ��ϵͳ", vbInformation, gstrSysName
    End If
    Call RecToLog(mrsSysInfo, "ϵͳ���", "��֤����ϵͳϵͳ��¼��")
End Sub

Private Sub cmdFunc_Click(Index As Integer)
    If cmdFunc(Index).ToolTipText <> "" Then
        MsgBox cmdFunc(Index).ToolTipText, vbInformation, gstrSysName
        Exit Sub
    End If
'    WriteTraceLog
    Call FuncControlExec(Index)
End Sub

Private Sub cmdkillProcess_Click()
    frmKillProcessManage.ShowMe ("0102")
End Sub

Private Sub cmdNotSel_Click()
    Call SetSelBeach
End Sub

Private Sub cmdSelAll_Click()
    Call SetSelBeach(True)
End Sub

Private Sub Form_Activate()
    If tbPage.Item(0).Selected Then
        Call VilidateUpgrade
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF5 Then 'ˢ�½���
        Call tbPage_SelectedChanged(tbPage.Item(IIf(tbPage.Item(0).Selected, 0, 1)))
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errH
    '����
    WriteTraceLog String(80, "/")
    WriteTraceLog String(4, "/") & "��������" & gstrServer
    WriteTraceLog String(4, "/") & "ʱ�䣺" & Format(CurrentDate, "yyyy-MM-dd HH:MM:SS")
    WriteTraceLog String(80, "/")
    '��ʼ������
    tbPage.Tag = "δ����"
    '��ʼ������
    tbPage.InsertItem 0, "��Ǩ����", picMain(0).hwnd, 0
    tbPage.InsertItem 1, "��Ǩ��ʷ", picMain(1).hwnd, 0
    tbPage.Tag = ""
    Call LoadSystems
    Call tbPage_SelectedChanged(tbPage.Item(0))
    cmdkillProcess.Visible = CheckAndAdjustMustTable("zlkillprocess")
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub Form_Resize()
    Dim i As Long
    On Error Resume Next
    tbPage.Height = Me.ScaleHeight - tbPage.Top + 15
    tbPage.Width = Me.ScaleWidth - tbPage.Left + 15
    For i = 0 To 1
        picMain(i).Left = 0
        picMain(i).Width = tbPage.Width - 60
        picMain(i).Height = tbPage.Height - picMain(i).Top
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsSysInfo = Nothing
    Set mrsSysUpFiles = Nothing
    mstrSysJobs = ""
    Set mrsMainPath = Nothing
    mblnLoadSysFiles = False
    mblnLastUpInfo = False
End Sub

Private Sub lblSel_Click()
    Dim strFolderName As String
    Dim strOldPath As String
    strFolderName = lblMainPath.Tag
    
    strFolderName = OpenFolder(Me, "ѡ��ϵͳ��װĿ¼")
    If strFolderName = "" Then Exit Sub
    If lblMainPath.Tag <> strFolderName Then
        lblMainPath.Tag = "": lblMainPath.Caption = "ϵͳ��װĿ¼��"
        Call GetAllSetup(strFolderName)
        Call optUpMode_Click(IIf(optUpMode(0).value, 0, 1))
    End If
End Sub

Private Sub optUpMode_Click(Index As Integer)
    With vsSysSel
        .Cell(flexcpText, .FixedRows, Col_Ŀ��汾, .Rows - 1, Col_�����) = ""
        .Cell(flexcpForeColor, .FixedRows, Col_Ŀ��汾, .Rows - 1, Col_�����) = &H80000008
        .Cell(flexcpChecked, .FixedRows, Col_Sel, .Rows - 1, Col_Sel) = True
        Call VilidateUpgrade
    End With
End Sub

Private Sub picMain_Resize(Index As Integer)
    Dim sngWidth As Long '��С���
    Dim i As Long
    
    If Index = 1 Then
        pgbState.Visible = False
        lblState.Visible = False
    End If
    On Error Resume Next
    sngWidth = picMain(0).ScaleWidth - 200
    If Index = 1 Then
        cboSys.Width = sngWidth - cboSys.Left - 300
        vsUpLog.Width = sngWidth - vsUpLog.Left - 300
        vsUpLog.Height = picMain(0).ScaleHeight - vsUpLog.Top - 100
    Else
        vsSysSel.Width = sngWidth - vsUpLog.Left - 90
        If vsSysSel.Top + vsSysSel.Rows * vsSysSel.RowHeightMin + cmdSelALl.Height + 200 < picMain(0).ScaleHeight Then
            vsSysSel.Height = vsSysSel.Rows * vsSysSel.RowHeightMin + 30
        Else
            vsSysSel.Height = IIf(vsSysSel.Rows < 13, vsSysSel.Rows, 12) * vsSysSel.RowHeightMin + 30
        End If
        lblExplain.Width = vsSysSel.Width
        lblExplain.Refresh
        'ϵͳ���Ʊ�ǩ��λ������
        Call SetCtrlPosOnLine(True, -1, lblExplain, 60, lblManage, 60, cmdFunc(FCT_�ͻ���), 90, lblState, 60, lblUpgrade, 60, lblMainPath, 90, lblUpMode, 90, lblConfigureFile, 30, vsSysSel)

        fraSplit(0).Left = -30: fraSplit(0).Width = lblManage.Left - fraSplit(0).Left
        Call SetCtrlPosOnLine(False, 0, lblManage, -1 * (lblManage.Width + fraSplit(0).Width), fraSplit(0), lblManage.Width, fraSplit(1))
        fraSplit(1).Width = picMain(0).ScaleWidth - fraSplit(1).Left + 100

        Call SetCtrlPosOnLine(False, 0, cmdFunc(FCT_�ͻ���), 120, cmdFunc(FCT_�û��˻�), 120, cmdFunc(FCT_��̨��ҵ), 120, cmdFunc(FCT_������), 720, cmdkillProcess)
        Call SetCtrlPosOnLine(False, 0, lblState, 0, pgbState)

        fraSplit(2).Left = -30: fraSplit(2).Width = lblUpgrade.Left - fraSplit(2).Left
        Call SetCtrlPosOnLine(False, 0, lblUpgrade, -1 * (lblUpgrade.Width + fraSplit(2).Width), fraSplit(2), lblUpgrade.Width, fraSplit(3))
        fraSplit(3).Width = picMain(0).ScaleWidth - fraSplit(3).Left + 100

        Call SetCtrlPosOnLine(False, 0, lblMainPath, 120, lblSel)
        Call SetCtrlPosOnLine(False, 0, lblUpMode, 120, fraUpMode, 120, pgbGetUpInfo)
        Call SetCtrlPosOnLine(True, 1, vsSysSel, 90, cmdExec)
        Call SetCtrlPosOnLine(True, -1, vsSysSel, 90, cmdSelALl)
        Call SetCtrlPosOnLine(False, 0, cmdSelALl, 60, cmdNotSel)
    End If
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tbPage.Tag = "" Then
        Call LoadData(Item.Index)
        picMain_Resize (Item.Index)
    End If
End Sub

Private Sub vsSysSel_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsSysSel
        .ComboList = ""
        .FocusRect = flexFocusLight
         .ToolTipText = ""
        If NewCol = Col_�����ļ� Then
             .ComboList = "..."
             .FocusRect = flexFocusSolid
        End If
    End With
End Sub

Private Sub vsSysSel_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strFile As String
    Dim strMainPath As String

    If Col = Col_�����ļ� Then
        With cdgPub
            .DialogTitle = "ѡ��Ӧ�ð�װ�����ļ�"
            If Trim(vsSysSel.TextMatrix(Row, Col_���)) = "" Then
                .Filter = "���������߽ű�(zlServer.Sql)|zlServer.Sql"
            Else
                .Filter = "Ӧ�ð�װ�����ļ�(zlSetup.ini)|zlSetup.ini"
                .Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
            End If
            strFile = IIf(Mid(vsSysSel.TextMatrix(Row, Col), 1, 1) = "$", lblMainPath.Tag & Mid(vsSysSel.TextMatrix(Row, Col), 2), vsSysSel.TextMatrix(Row, Col))
            If gobjFile.FileExists(strFile) Then
                .InitDir = gobjFile.GetParentFolderName(strFile)
                .Filename = gobjFile.GetFileName(strFile)
            Else
                .InitDir = "": .Filename = ""
                If vsSysSel.Cell(flexcpData, Row, Col) <> "" Then
                    If gobjFile.FolderExists(gobjFile.GetParentFolderName(vsSysSel.Cell(flexcpData, Row, Col))) Then
                        .InitDir = gobjFile.GetParentFolderName(vsSysSel.Cell(flexcpData, Row, Col))
                    End If
                End If
            End If
            On Error Resume Next
            .CancelError = True
            .ShowOpen
            err.Clear: On Error GoTo errH
            If .Filename = gobjFile.GetFileName(strFile) Then .Filename = ""
            If .Filename <> "" And .Filename <> "zlSetup.ini" And .Filename <> "zlServer.Sql" Then
                If .Filename <> vsSysSel.Cell(flexcpData, Row, Col) Then
                    '�����ļ��ı䣬��������ļ�
                    If CheckInitFile(Val(vsSysSel.TextMatrix(Row, Col_���)), .Filename) Then
                        vsSysSel.TextMatrix(Row, Col) = .Filename
                         vsSysSel.Cell(flexcpData, Row, Col) = .Filename
                        Call ReSetMainPath(Row)
                        Call VilidateUpgrade(Row)
                    End If
                End If
            End If
            On Error GoTo 0
        End With
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub vsSysSel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (vsSysSel.MouseCol = Col_Ŀ��汾 Or vsSysSel.MouseCol = Col_�����) And vsSysSel.MouseRow >= vsSysSel.FixedRows Then
        If vsSysSel.TextMatrix(vsSysSel.MouseRow, Col_�����) <> "" Then
            vsSysSel.ToolTipText = vsSysSel.TextMatrix(vsSysSel.MouseRow, Col_�����)
        Else
            vsSysSel.ToolTipText = ""
        End If
    Else
        vsSysSel.ToolTipText = ""
    End If
End Sub

Private Sub vsSysSel_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = Col_Sel And Row > vsSysSel.FixedRows Or Col = Col_�����ļ�) Then Cancel = True
End Sub

Private Sub vsUpLog_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsUpLog
        If NewRow >= .FixedRows Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, NewCol)
        End If
    End With
End Sub

'===========================================================
'����
'===========================================================
Private Sub LoadSystems()
'���ܣ�����Ӧ��ϵͳ
'������intPageIndex=0����Ǩҳϵͳ��ӣ�intPageIndex=1,��Ǩ��ʷҳϵͳ���
    Dim strSQL As String, rsSys As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strVer As String
    Dim i As Long
    On Error GoTo errH
    '��ȡ�����߰汾��
    strVer = GetToolsVersion
    '���ӹ����������Ҫ�ǽ���ϵͳ����ǰ��
    strSQL = "Select ��� ϵͳ���, ���� ϵͳ����, �汾�� ϵͳ�汾��, ������ ϵͳ������, �����, ������װ From Zlsystems where Upper(������)=[1] Order by Nvl(�����,0),���"
    Set rsSys = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ��װϵͳ", gstrUserName)
    With rsSys
        '��ӹ�������ʷ��¼�鿴��
        cboSys.Clear
        cboSys.addItem String(5, " ") & RPAD("������������", 18) & " v" & VerPAD(strVer)
        cboSys.ItemData(cboSys.NewIndex) = -1
        Do While Not .EOF
            If Val(Split(!ϵͳ�汾��, ".")(0)) > 9 Then
                    cboSys.addItem Lpad(!ϵͳ���, 4) & "-" & RPAD(!ϵͳ���� & "", 18) & " v" & VerPAD(!ϵͳ�汾�� & "")
                    cboSys.ItemData(cboSys.NewIndex) = !ϵͳ���
                    If cboSys.ListIndex = -1 And UCase(!ϵͳ������ & "") = UCase(gstrUserName) Then
                        cboSys.ListIndex = cboSys.NewIndex
                    End If
            End If
            .MoveNext
        Loop
        If cboSys.ListIndex = -1 Then cboSys.ListIndex = 0
    End With
    If rsSys.RecordCount <> 0 Then rsSys.MoveFirst
    '��д�Ѱ�װϵͳ�嵥
    With vsSysSel
        'Ŀ��汾�����հ汾Ϊϵͳ��������ʱ�ı�����ǨĿ���Լ�����Ŀ��
        Set mrsSysInfo = CopyNewRec(rsSys, True, "ϵͳ���,ϵͳ����,ϵͳ�汾��,ϵͳ������,�����,������װ", Array("Sort", adInteger, 2, 0, "����", adInteger, 1, 0, "�����ļ�", adVarChar, 2000, Empty, _
                                                                                       "Ŀ��汾", adVarChar, 20, Empty, "��ǰĿ��汾", adVarChar, 20, Empty, "���հ汾", adVarChar, 20, Empty, _
                                                                                        "��Ǩ���", adInteger, 1, 0, "��ֹ��Ϣ", adVarChar, 2000, Empty, "������", adInteger, 1, 0, "�����", adVarChar, 2000, Empty, _
                                                                                        "��ǰ��Ǩ���", adInteger, 1, 0, "��ǰ��ֹ��Ϣ", adVarChar, 2000, Empty, "����ǰ����", adInteger, 1, 0, "��ǰ�����", adVarChar, 2000, Empty))
        .Rows = .FixedRows
        '��ȡ�����߰汾��
        strVer = GetToolsVersion
        mrsSysInfo.AddNew Array("ϵͳ���", "ϵͳ����", "ϵͳ�汾��", "ϵͳ������", "�����", "������װ", "Sort", "�����ļ�", "������", "����ǰ����", "����"), _
                                        Array(0, "������", strVer, "ZLTOOLS", Null, 1, .Rows, Null, 1, 1, 1)
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, Col_Sel) = IIf(strVer & "" = "", 0, 1)
        .TextMatrix(.Rows - 1, Col_���) = ""
        .TextMatrix(.Rows - 1, Col_����) = "������������"
        .TextMatrix(.Rows - 1, Col_��ǰ�汾) = VerPAD(strVer & "")
        .TextMatrix(.Rows - 1, Col_�����) = ""
        .Cell(flexcpForeColor, .Rows - 1, Col_Sel, .Rows - 1, .Cols - 1) = IIf(strVer & "" = "", vbRed, vbBlue)
        Do While Not rsSys.EOF
            If Val(Split(rsSys!ϵͳ�汾��, ".")(0)) > 9 Then
                mrsSysInfo.AddNew Array("ϵͳ���", "ϵͳ����", "ϵͳ�汾��", "ϵͳ������", "�����", "������װ", "Sort", "�����ļ�", "������", "����ǰ����", "����"), _
                                                Array(rsSys!ϵͳ���, rsSys!ϵͳ����, rsSys!ϵͳ�汾��, rsSys!ϵͳ������, rsSys!�����, rsSys!������װ, .Rows, Null, 1, 1, 1)
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, Col_Sel) = 1
                .TextMatrix(.Rows - 1, Col_���) = rsSys!ϵͳ��� & ""
                .Cell(flexcpData, .Rows - 1, Col_���) = Val(rsSys!����� & "")
                .TextMatrix(.Rows - 1, Col_����) = rsSys!ϵͳ���� & ""
                .TextMatrix(.Rows - 1, Col_��ǰ�汾) = VerPAD(rsSys!ϵͳ�汾�� & "")
                .TextMatrix(.Rows - 1, Col_�����) = ""
            End If
            rsSys.MoveNext
        Loop
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, Col_����) <> 0 Then
                mrsSysInfo.Filter = "ϵͳ���=" & .Cell(flexcpData, i, Col_����)
                .RowData(i) = Val(mrsSysInfo!��� & "")
            End If
        Next
        Call GetLastUpgrade
    End With
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub GetLastUpgrade()
'���ܣ���ȡ�ϴ���Ǩ��Ϣ
    Dim rsUpgrade As ADODB.Recordset
    Dim strSQL As String, strFilter As String
    Dim lngSys As Long
    Dim i As Long
    
    On Error GoTo errH
    mblnLastUpInfo = False
    '���ZLUPGRADE�����ֶΡ���ǰִ�С�
    If Not CheckAndAdjustMustTable("ZLUPGRADE", "��ǰִ��", True) Then
        Exit Sub
    End If
    If cboSys.ListCount > 1 Then
        '����ZLBAKSPACES
        If Not CheckAndAdjustMustTable("ZLBAKSPACES", , True) Then
            Exit Sub
        End If
        '����ZLBAKTABLES
        If Not CheckAndAdjustMustTable("ZLBAKTABLES", , True) Then
            Exit Sub
        End If
    End If
    mblnLastUpInfo = True
    '��ȡ����ϵͳ�ϴ���Ǩ�Լ��ϴ���ǰ��Ǩ��Ϣ
    strSQL = "Select Nvl(ϵͳ,0) ϵͳ��� , ��ǰִ��, ��ֹ���, ��Ǩ���, ����汾" & vbNewLine & _
                    "From (Select ϵͳ, ��ǰִ��, ��Ǩʱ��, ��ֹ���, ��Ǩ���, ����汾, Max(��Ǩʱ��) Over(Partition By ϵͳ, Decode(��ǰִ��, Null, -1, 0)) ��ǰʱ��" & vbNewLine & _
                    "       From Zlupgrade) a" & vbNewLine & _
                    "Where A.��Ǩʱ�� = A.��ǰʱ��" & vbNewLine & _
                    "Order By ϵͳ"
    Set rsUpgrade = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ�ϴ���Ǩ��Ϣ")
    
    For i = vsSysSel.FixedRows To vsSysSel.Rows - 1
        lngSys = Val(vsSysSel.TextMatrix(i, Col_���))
        strFilter = "ϵͳ��� = " & lngSys
        mrsSysInfo.Filter = strFilter
        'ϵͳ�ϴ�ִ����Ǩ��Ϣ
        rsUpgrade.Filter = strFilter & " And  ��ǰִ��=Null"
        If Not rsUpgrade.EOF Then
            mrsSysInfo.Update Array("��Ǩ���", "��ֹ��Ϣ"), Array(rsUpgrade!��Ǩ���, FormatUpgradeBreak(lngSys, rsUpgrade!����汾 & "", rsUpgrade!��ֹ��� & ""))
            'ϵͳ���һ�������������ɹ������ܽ�����ǰִ��
            If Val(rsUpgrade!��Ǩ��� & "") = 1 Then
                mrsSysInfo.Update Array("����ǰ����", "��ǰ�����"), Array(0, "ϵͳ���һ�������������ɹ������ܽ�����ǰִ�У�")
            End If
        Else
            mrsSysInfo.Update Array("��Ǩ���", "��ֹ��Ϣ"), Array(0, FormatUpgradeBreak(lngSys, mrsSysInfo!ϵͳ�汾�� & ""))
        End If
        'ϵͳ�ϴ�ִ����ǰ��Ǩ��Ϣ
        rsUpgrade.Filter = strFilter & " And ��ǰִ��<>Null"
        If Not rsUpgrade.EOF Then
            mrsSysInfo.Update Array("��ǰ��Ǩ���", "��ǰ��ֹ��Ϣ"), Array(rsUpgrade!��Ǩ���, FormatUpgradeBreak(lngSys, rsUpgrade!����汾 & "", rsUpgrade!��ֹ��� & ""))
        Else
            mrsSysInfo.Update Array("��ǰ��Ǩ���", "��ǰ��ֹ��Ϣ"), Array(0, FormatUpgradeBreak(lngSys, mrsSysInfo!ϵͳ�汾�� & ""))
        End If
    Next
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub LoadData(ByVal intPageIdx As Integer)
'���ܣ����ݼ���
'    intPageIdx=ҳ��������1-��Ǩҳ��,0-��־����
    Dim blnUp As Boolean
    
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strCtrlInfo As String
    Dim arrInfo As Variant, arrTmp As Variant, i As Long
    Dim lngSys As Long
    Dim intIndex As Long, strCaption As String, strTag As String
    
    On Error GoTo errH
    If intPageIdx = 1 Then
        lngSys = cboSys.ItemData(cboSys.ListIndex)
        If lngSys = Val(cboSys.Tag) Then Exit Sub
        cboSys.Tag = lngSys
        strSQL = "Select * From zlUpgrade Where " & IIf(lngSys = -1, "ϵͳ Is Null ", "ϵͳ=[1] ") & " Order by ��Ǩʱ��"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ��Ǩ��ʷ", lngSys)
        With vsUpLog
            .Rows = 1
            On Error Resume Next
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, Col_���) = .Rows - 1
                .TextMatrix(.Rows - 1, Col_��Ǩʱ��) = Format(rsTmp!��Ǩʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(.Rows - 1, Col_ԭʼ�汾) = VerPAD(rsTmp!ԭʼ�汾 & "")
                .TextMatrix(.Rows - 1, Col_Ԥ��Ŀ��) = VerPAD(rsTmp!Ŀ��汾 & "")
                .TextMatrix(.Rows - 1, Col_����汾) = VerPAD(rsTmp!����汾 & "")
                .TextMatrix(.Rows - 1, Col_��Ǩ���) = IIf(Nvl(rsTmp!��Ǩ���, 0) = 0, "�������", "��;��ֹ")
                '����û����ǰִ����һ��
                .TextMatrix(.Rows - 1, Col_��ǰִ��) = rsTmp!��ǰִ�� & ""
                If rsTmp!��ǰִ�� & "" <> "" Then
                    .TextMatrix(.Rows - 1, Col_��ǰִ��) = "��"
                End If
                If Nvl(rsTmp!��Ǩ���, 0) <> 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                End If
                rsTmp.MoveNext
            Loop
            err.Clear: On Error GoTo errH
            If .Rows > 1 Then
                .Row = .Rows - 1
                .ShowCell .Row, .Col
            End If
        End With
    Else
        '���й��������:
        '1��û�а�װ����Ӧ��ϵͳ
        '2����װ��ϵͳ�������������û���¼��SYS�������û���¼����ϵͳ������
        '��ʱ��Ҫ�����´���
        If vsSysSel.Tag = "" Then
            lblMainPath.Tag = App.Path
            lblMainPath.Caption = "ϵͳ��װĿ¼��" & App.Path
            Call GetAllSetup
            vsSysSel.Tag = "�Ѿ�����"
        End If
        '�ṹ��֤,�����Ѿ���������һ�μ��Ľ������˲��õ����������⣬ÿ�����õĶ������Ѿ���������
        '���Zlupgradeconfig��
        If Not CheckAndAdjustMustTable("Zlupgradeconfig", , True) Then
            For i = 0 To 3
                cmdFunc(i).ToolTipText = "��Zlupgradeconfigȱʧ���޷�ִ�иù��ܣ�"
            Next
            Exit Sub
        End If
        '���Zlclients���롰ϵͳ�������á��ֶ�
        If Not CheckAndAdjustMustTable("Zlclients", "ϵͳ��������", True) Then
            cmdFunc(FCT_�ͻ���).ToolTipText = "��Zlclientsȱʧ���߱�Zlclients���ֶ�""ϵͳ��������""ȱʧ���޷�ִ�иù��ܣ�"
        End If
        '���ZLAutoJob���롰ϵͳ����ͣ�á��ֶ�
        If Not CheckAndAdjustMustTable("ZLAutoJobs", "ϵͳ����ͣ��", True) Then
            cmdFunc(FCT_��̨��ҵ).ToolTipText = "��ZLAutoJobȱʧ���߱�ZLAutoJob���ֶ�""ϵͳ����ͣ��""ȱʧ���޷�ִ�иù��ܣ�"
        End If
        If cboSys.ListCount > 1 Then
            '����ϻ���Ա����롰ϵͳ�����������ֶ�
            If Not CheckAndAdjustMustTable("�ϻ���Ա��", "ϵͳ��������", True, gstrUserName) Then
                cmdFunc(FCT_�û��˻�).ToolTipText = "���ϻ���Ա��ȱʧ���߱��ϻ���Ա����ֶ�""ϵͳ��������""ȱʧ���޷�ִ�иù��ܣ�"
            End If
        Else
            cmdFunc(FCT_�û��˻�).ToolTipText = "����Ӧ�������ߵ�¼������ִ�иù��ܣ�"
        End If
        '���ZLTriggers��
        If Not CheckAndAdjustMustTable("ZLTriggers", , True) Then
            cmdFunc(FCT_������).ToolTipText = "��ZLTriggersȱʧ���޷�ִ�иù��ܣ�"
        End If
        '��ȡϵͳ��Ǩ׼��������ť״̬�����ð�ť״̬
        strSQL = "Select ��Ŀ, ����" & vbNewLine & _
                        "From Zlupgradeconfig" & vbNewLine & _
                        "Where ��Ŀ In ('�ͻ���״̬', '�û�״̬', '��̨��ҵ״̬', '������״̬', '���õ�ϵͳ����')"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡϵͳ��Ǩ׼��������ť״̬")
        Do While Not rsTmp.EOF
            intIndex = -1
            Select Case rsTmp!��Ŀ & ""
                Case "�ͻ���״̬"
                    strCaption = IIf(Val(rsTmp!���� & "") = 1, "���ò��жϿͻ���", "���ÿͻ���"): intIndex = FCT_�ͻ���
                Case "�û�״̬"
                    strCaption = IIf(Val(rsTmp!���� & "") = 1, "�����û��ʺ�", "�ָ��û��ʺ�"): intIndex = FCT_�û��˻�
                Case "��̨��ҵ״̬"
                    strCaption = IIf(Val(rsTmp!���� & "") = 1, "���ú�̨��ҵ", "���ú�̨��ҵ"): intIndex = FCT_��̨��ҵ
                Case "������״̬"
                    strCaption = IIf(Val(rsTmp!���� & "") = 1, "���ô�����", "���ô�����"): intIndex = FCT_������
                Case "���õ�ϵͳ����"
                    mstrSysJobs = rsTmp!���� & ""
            End Select
            strTag = Val(rsTmp!���� & "") & ""
            If intIndex <> -1 Then
                cmdFunc(intIndex).Caption = strCaption: cmdFunc(intIndex).Tag = strTag
            End If
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub ReSetMainPath(Optional ByVal lngRow As Long = -1)
'���ܣ���·��û�б�ʹ�����Զ�������·����ʹ����·����·���Զ��޸�Ϊ��дģʽ
'        :lngRow=��ǰ�޸���
    Dim blnRest As Boolean '�Ƿ�����·��
    Dim i As Long, lngTmpRow As Long
    Dim strMainPath As String, strMainPathOld As String
    Dim blnNotHave As Boolean
    
    On Error GoTo errH
    With vsSysSel
        blnRest = True
        If lblMainPath.Tag <> "" Then
            If lngRow >= .FixedRows Then
                If .TextMatrix(lngRow, Col_�����ļ�) = "" Then lngRow = -1
            End If
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, Col_�����ļ�) <> "" Then
                    If UCase(Mid(.TextMatrix(i, Col_�����ļ�), 1, Len(lblMainPath.Tag) + 1)) = UCase(lblMainPath.Tag) & "\" Then
                        .TextMatrix(i, Col_�����ļ�) = "$" & Mid(.TextMatrix(i, Col_�����ļ�), Len(lblMainPath.Tag) + 1)
                        blnRest = False
                    ElseIf Mid(.TextMatrix(i, Col_�����ļ�), 1, 1) = "$" Then
                        blnRest = False
                    End If
                    If lngTmpRow = 0 Then lngTmpRow = i
                End If
            Next
        End If
        If blnRest Then
            On Error Resume Next
            If lngRow >= lngTmpRow Then
                strMainPath = gobjFile.GetFile(.Cell(flexcpData, lngRow, Col_�����ļ�)).ParentFolder.ParentFolder.ParentFolder
            Else
                strMainPath = gobjFile.GetFile(.Cell(flexcpData, lngTmpRow, Col_�����ļ�)).ParentFolder.ParentFolder
            End If
            If err.Number <> 0 Then
                err.Clear
            End If
            On Error GoTo errH
            If strMainPath <> "" Then
                '������·��
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, Col_�����ļ�) <> "" Then
                        If UCase(Mid(.TextMatrix(i, Col_�����ļ�), 1, Len(strMainPath) + 1)) = UCase(strMainPath) & "\" Then 'Ӧ�ó���װ·�����ã��򲻸ı�
                            .TextMatrix(i, Col_�����ļ�) = "$" & Mid(.TextMatrix(i, Col_�����ļ�), Len(strMainPath) + 1)
                        End If
                    End If
                Next
                lblMainPath.Tag = strMainPath
                lblMainPath.Caption = "ϵͳ��װĿ¼��" & strMainPath
            End If
        End If
    End With
    Call SetCtrlPosOnLine(False, 0, lblMainPath, 120, lblSel)
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub GetAllSetup(Optional ByVal strMainPath As String)
'���ܣ���ȡZLSOFT�������ϵͳ��װ�����ļ�
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim rsMainPath As ADODB.Recordset
    Dim strPath As String
    Dim strFile As String
    Dim i As Integer, blnAdd As Boolean
    
    On Error GoTo errH
    '����ϴ�����
    vsSysSel.Cell(flexcpText, vsSysSel.FixedRows, Col_�����ļ�, vsSysSel.Rows - 1, Col_��ǰ�汾 - 1) = ""
    vsSysSel.Cell(flexcpData, vsSysSel.FixedRows, Col_�����ļ�, vsSysSel.Rows - 1, Col_��ǰ�汾 - 1) = ""
    vsSysSel.Cell(flexcpText, vsSysSel.FixedRows, Col_��ǰ�汾 + 1, vsSysSel.Rows - 1, vsSysSel.Cols - 1) = ""
    vsSysSel.Cell(flexcpData, vsSysSel.FixedRows, Col_��ǰ�汾 + 1, vsSysSel.Rows - 1, vsSysSel.Cols - 1) = ""
    mblnLoadSysFiles = False
    '��ȡ��װ�����ļ����ѡ��Ŀ¼
    If mrsMainPath Is Nothing Or strMainPath <> "" Then
        Set mrsMainPath = CopyNewRec(Nothing, True, , Array("���", adInteger, 3, 0, "ϵͳ���", adInteger, 5, 0, "·��", adVarChar, 2000, Empty))
        On Error Resume Next
        '0����ִ����Ŀ¼�������Ŀ¼����
        If strMainPath <> "" Then
            mrsMainPath.AddNew Array("���", "ϵͳ���", "·��"), Array(1, 0, UCase(strMainPath))
        End If
        '1������ͨ��ͨ��ע���ȷ��,ע�������������ϵͳ�����װϵͳ�ܻ����ע����Ϣ
        strPath = GetSetting("ZLSOFT", "����ȫ��", "����·��")
        strPath = gobjFile.GetFile(strPath).ParentFolder
        If err.Number = 0 Then
            mrsMainPath.Filter = "·��='" & UCase(strPath) & "'"
            If mrsMainPath.EOF Then mrsMainPath.AddNew Array("���", "ϵͳ���", "·��"), Array(2, 0, UCase(strPath))
        Else
            err.Clear
        End If
        'ͨ��ϵͳĿ¼��ȡ
        strPath = gobjFile.GetFolder(Mid(gobjFile.GetSpecialFolder(WindowsFolder), 1, 1) & ":\APPSOFT")
        If err.Number = 0 Then
            mrsMainPath.Filter = "·��='" & UCase(strPath) & "'"
            If mrsMainPath.EOF Then mrsMainPath.AddNew Array("���", "ϵͳ���", "·��"), Array(3, 0, UCase(strPath))
        Else
            err.Clear
        End If
        '2������10�汾ϵͳ�İ�װ�����ļ�ȷ��
        '3��ͨ������10�汾ϵͳ�����������ļ�ȷ��
        strSQL = "Select A.ϵͳ ϵͳ���, A.����, A.�ļ��� From Zlsysfiles a Where  A.���� in(1,2) Order By ϵͳ,����"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡϵͳ������װ�������ļ�")
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!���� & "") = 1 Then
                strPath = gobjFile.GetFile(rsTmp!�ļ��� & "").ParentFolder.ParentFolder.ParentFolder
                strFile = rsTmp!�ļ��� & ""
            Else
                strPath = gobjFile.GetFile(rsTmp!�ļ��� & "").ParentFolder.ParentFolder.ParentFolder.ParentFolder
                strFile = gobjFile.GetFile(rsTmp!�ļ��� & "").ParentFolder.ParentFolder.ParentFolder & "\Ӧ�ýű�\ZLSETUP.INI"
            End If
            If err.Number = 0 Then
                mrsMainPath.Filter = "·��='" & UCase(strPath) & "' And ϵͳ���=0"
                If mrsMainPath.EOF Then mrsMainPath.AddNew Array("���", "ϵͳ���", "·��"), Array(i + 3, 0, UCase(strPath))
                If Not gobjFile.FileExists(strFile) Then strFile = ""
            Else
                err.Clear
                strFile = ""
            End If
            If strFile <> "" Then
                mrsMainPath.Filter = "·��='" & UCase(strFile) & "' And ϵͳ���=" & rsTmp!ϵͳ
                If mrsMainPath.EOF Then mrsMainPath.AddNew Array("���", "ϵͳ���", "·��"), Array(i + 4, rsTmp!ϵͳ���, UCase(strFile))
            End If
            rsTmp.MoveNext
        Next
    End If
    mrsMainPath.Filter = "ϵͳ���<>0"
    mblnLoadSysFiles = mrsMainPath.RecordCount = 0 'û�ж�ȡ��ZLSysFiles����Ĭ���Ѿ�����
    mrsMainPath.Filter = "ϵͳ���=0"
    mrsMainPath.Sort = "���,·��"
    If mrsMainPath.RecordCount <> 0 Then
        blnAdd = strMainPath = ""
        For i = 0 To mrsMainPath.RecordCount - 1
            If mrsMainPath!·�� & "" <> "" Then
                If GetSetupInit(mrsMainPath!·�� & "", blnAdd) Then Exit For
                If blnAdd Then blnAdd = Not blnAdd
            End If
            mrsMainPath.MoveNext
        Next
        '����·�����������ַ���ʶ��û��ʹ����·�������Զ�����
        Call ReSetMainPath
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function GetSetupInit(Optional ByVal strMainPath As String, Optional ByRef blnAdd As Boolean) As Boolean
'���ܣ���ȡ����ϵͳ�İ�װ�����ļ�
'������strMainPath="",ͨ��ϵͳ�ļ�ZLSysFiles��ȡ�ļ���<>""��ͨ��+·����ȡ�ļ�
'           blnAdd=�Ƿ�ֻ��ȡδ��ȡ��ϵͳ�������ļ�
    Dim strPath As String, lngCurSys As Long
    Dim strFile As String
    Dim blnGet As Boolean, blnAllGet As Boolean, blnToolsGet As Boolean, blnSysFileGet As Boolean
    Dim strTmp As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    
    With vsSysSel
        '�Զ���ȡʱ���������ϴα����ZLSysFiles��ȡ
        If blnAdd And Not mblnLoadSysFiles Then Call LoadSysFiles
        '�Զ���ȡ����Ŀ¼��ȡ
        blnAllGet = True
        For i = .FixedRows To .Rows - 1
            lngCurSys = Val(.TextMatrix(i, Col_���))
            If blnAdd And .TextMatrix(i, Col_�����ļ�) = "" Or Not blnAdd Then
                If lngCurSys = 0 Then
                    strTmp = "\TOOLS\ZLSERVER.SQL"
                    strFile = strMainPath & strTmp
                Else
                    strTmp = "\" & Decode(lngCurSys \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                                6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                                23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                                25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\Ӧ�ýű�\ZLSETUP.INI"
                    strFile = strMainPath & strTmp
                End If
                If gobjFile.FileExists(strFile) Then
                    If CheckInitFile(lngCurSys, strFile, True) Then
                        .Cell(flexcpData, i, Col_�����ļ�) = gobjFile.GetFile(strFile).Path
                        .TextMatrix(i, Col_�����ļ�) = .Cell(flexcpData, i, Col_�����ļ�)
                        blnGet = True
                    End If
                End If
                If .TextMatrix(i, Col_�����ļ�) = "" Then blnAllGet = False
            End If
            '�Ƿ��ȡ�˹����������ļ�
            If .TextMatrix(i, Col_�����ļ�) <> "" And lngCurSys = 0 Then
                blnToolsGet = True
            End If
        Next
        '�ֹ�ָ����Ŀ¼�����ZLSYsFiles�е������ļ�
        If Not blnAdd And Not mblnLoadSysFiles Then
            blnSysFileGet = LoadSysFiles
            blnAllGet = blnSysFileGet And blnToolsGet
        End If
        If Not blnAdd And Not blnGet Then
            MsgBox "����Ŀ¼" & strMainPath & "��δ�ҵ��κ�ϵͳ��װ�����ļ���ϵͳ���Զ���ȡ��װ�����ļ���"
        Else
            '������Ŀ¼
            If blnGet And lblMainPath.Tag = "" Then
                lblMainPath.Tag = gobjFile.GetFolder(strMainPath).Path
                lblMainPath.Caption = "ϵͳ��װĿ¼��" & lblMainPath.Tag
            End If
        End If
    End With

    GetSetupInit = blnAllGet
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Function

Private Function LoadSysFiles() As Boolean
'���ܣ�����ZLSysFiles�еļ�¼�İ�װ�����ļ�
    Dim blnAllGet As Boolean, i As Long
    Dim lngCurSys As Long
    
    On Error GoTo errH
    With vsSysSel
        blnAllGet = True
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Col_�����ļ�) = "" Then
                lngCurSys = Val(.TextMatrix(i, Col_���))
                If lngCurSys <> 0 Then
                    mrsMainPath.Filter = "ϵͳ���=" & lngCurSys
                    mrsMainPath.Sort = "���"
                    Do While Not mrsMainPath.EOF
                        If gobjFile.FileExists(mrsMainPath!·�� & "") Then
                            If CheckInitFile(lngCurSys, mrsMainPath!·�� & "", True) Then
                                .Cell(flexcpData, i, Col_�����ļ�) = gobjFile.GetFile(mrsMainPath!·�� & "").Path
                                .TextMatrix(i, Col_�����ļ�) = gobjFile.GetFile(mrsMainPath!·�� & "").Path
                                Exit Do
                            End If
                        End If
                        mrsMainPath.MoveNext
                    Loop
                    If .TextMatrix(i, Col_�����ļ�) = "" Then blnAllGet = False
                End If
            End If
        Next
    End With
    mblnLoadSysFiles = True
    LoadSysFiles = blnAllGet
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub FuncControlExec(ByVal fctTmp As FuncCtrlType)
'���ܣ�ִ������ǰ�ù���
'������fctTmp=��������
    Dim strInfoName As String
    Dim blnLock As Boolean '�Ƿ���������
    Dim strLog As String
    Dim rsSQL As ADODB.Recordset
    Dim strStep As String
    Dim cnDBA As ADODB.Connection
    '��ȡ������־����
    '�����洴����ʹ�����رգ���ֹһ̨���Կ�����������໥ռ����־�ļ�
    strLog = GetLogPath(LT_ϵͳ����)
    If mobjOprateLog Is Nothing Then
        If gobjFile.FileExists(strLog) Then
            Set mobjOprateLog = gobjFile.OpenTextFile(strLog, ForAppending)
        Else
            Set mobjOprateLog = gobjFile.CreateTextFile(strLog, True)
        End If
    End If
    On Error GoTo errH
    '��ȡ��Ϣ��
    strInfoName = Decode(fctTmp, FCT_�ͻ���, "�ͻ���״̬", FCT_�û��˻�, "�û�״̬", FCT_��̨��ҵ, "��̨��ҵ״̬", FCT_������, "������״̬")
    '��ť״̬��֤����ֹ�ظ�������״̬δ������֤������ʾ�Ƿ��������
    If Not FuncControlStateVali(fctTmp, strInfoName) Then Exit Sub
    blnLock = Val(cmdFunc(fctTmp).Tag) = 1 '��ǰ������״̬������ö���
    '��ȡִ��SQL
    Set rsSQL = FuncControlSQL(fctTmp, blnLock, strInfoName)
    rsSQL.Filter = "DBA=1"
    If Not rsSQL.EOF Then
        Set cnDBA = GetConnection("System")
        If cnDBA Is Nothing Then Exit Sub
    End If
    mobjOprateLog.WriteLine String(120, "=")
    mobjOprateLog.WriteLine strInfoName & IIf(blnLock, " ����-->����", " ����-->����")
    mobjOprateLog.WriteLine "ִ��ʱ�䣺" & Format(CurrentDate, "yyyy-MM-dd HH:mm:ss")
    mobjOprateLog.WriteLine String(120, "=")
    'ִ��SQL
    rsSQL.Filter = ""
    rsSQL.Sort = "���"
    On Error Resume Next
    Do While Not rsSQL.EOF
        Call SetProgressState(Val(rsSQL!���� & ""), rsSQL!���� & "")
        mobjOprateLog.WriteLine "ִ��SQL:" & rsSQL!����
        If rsSQL!DBA = 1 Then
            cnDBA.Execute rsSQL!����
        Else
            gcnOracle.Execute rsSQL!����
        End If
        If err.Number <> 0 Then
            mobjOprateLog.WriteLine "ִ�н��:" & err.Description
            err.Clear
        Else
            mobjOprateLog.WriteLine "ִ�н��:ִ�гɹ�" & err.Description
        End If
        rsSQL.MoveNext
    Loop
    On Error GoTo errH
    '�����ǣ���ֹ�ٴδ����Լ�����,��ˢ�½��棬���ð�ť״̬
    Call LoadData(0)
    '�������
    Call SetProgressState(101)
     '�����洴����ʹ�����رգ���ֹһ̨���Կ�����������໥ռ����־�ļ�
    mobjOprateLog.Close
    Set mobjOprateLog = Nothing
    Me.Refresh
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function FuncControlStateVali(ByVal fctTmp As FuncCtrlType, ByVal strInfoName As String) As Boolean
'���ܣ����ܰ�ť״̬��֤
'������fctTmp =�ؼ�����
'         strInfoName=�ռ�ʹ�õ���Ϣ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String, blnCando As Boolean
    '���ý���
    On Error GoTo errH
    strSQL = "Select ��Ŀ, ���� From Zlupgradeconfig Where ��Ŀ =[1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, strInfoName)
    blnCando = True
    If rsTmp.EOF Then
        If Val(cmdFunc(fctTmp).Tag) <> 1 Then
            blnCando = False
        End If
        On Error Resume Next
        strSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('" & strInfoName & "',1)"
        gcnOracle.Execute strSQL
        If err.Number <> 0 Then err.Clear
        On Error GoTo errH
    Else
        If Val(rsTmp!���� & "") <> Val(cmdFunc(fctTmp).Tag) Then
            blnCando = False
        End If
    End If
    If Not blnCando Then
        MsgBox "����""" & cmdFunc(fctTmp).Caption & """�Ѿ�ִ�У�", vbInformation, gstrSysName
        Call LoadData(0) 'ˢ�½���
        Exit Function
    End If
    FuncControlStateVali = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Function FuncControlSQL(ByVal fctTmp As FuncCtrlType, ByVal blnLock As Boolean, ByVal strInfoName As String) As ADODB.Recordset
'���ܣ���������ִ�еĲ����Լ�SQL
'������fctTmp =�ؼ�����
'         blnLock=�Ƿ���ö���
    Dim strSQL As String
    Dim strAdjustSQL As String
    Dim rsObjectSQL As ADODB.Recordset, rsSysScheduler As ADODB.Recordset
    Dim rsSQLStep As ADODB.Recordset '����SQL�Լ�����
    Dim strTmp As String, rsTmp As ADODB.Recordset
    Dim lngSort As Long, lngAll As Long
    Dim arrTmp As Variant, i As Long
  
  
    On Error GoTo errH
    '���=SQLִ��˳���Լ�����˳��
    '����=0��SQL,1:�����Լ�����
    Set rsSQLStep = CopyNewRec(Nothing, True, , Array("���", adInteger, 5, 0, "����", adVarChar, 100, Empty, "����", adDouble, 5, 0, "����", adVarChar, 2000, Empty, "DBA", adInteger, Empty, Empty))
    Select Case fctTmp
        Case FCT_�ͻ���
            '����SQL�Լ�����
            lngAll = lngAll + 2
            If blnLock Then
                Call KillSessions
            Else
                '���ÿͻ���
                strAdjustSQL = "Update Zlclients Set ��ֹʹ�� = 0, ϵͳ�������� = Null Where Nvl(ϵͳ��������, 0) = 1"
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, IIf(blnLock, "���ÿͻ���", "���ÿͻ���"), (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
        Case FCT_�û��˻�
            If blnLock Then
                '����������û�
                strAdjustSQL = "Update " & gstrUserName & ".�ϻ���Ա�� b" & vbNewLine & _
                                    "Set ϵͳ�������� = 1" & vbNewLine & _
                                    "Where Exists (Select 1 From Dba_Users a Where Account_Status = 'OPEN' And A.Username = Upper(B.�û���)) And Upper(B.�û���)<>'" & UCase(gstrUserName) & "'"
                '�����û��˻�
                strSQL = "Select 'alter user ' || �û��� || ' account lock '  SQL" & vbNewLine & _
                            "From " & gstrUserName & ".�ϻ���Ա�� b" & vbNewLine & _
                            "Where Exists (Select 1 From Dba_Users a Where Account_Status = 'OPEN' And A.Username = Upper(B.�û���)) And Upper(B.�û���)<>[1] " & vbNewLine & _
                            "Order By �û���"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, UCase(gstrUserName))
            Else
                '�����û��˻�
                strSQL = "Select 'alter user ' || �û��� || ' account unlock ' SQL From " & gstrUserName & ".�ϻ���Ա�� Where Nvl(ϵͳ��������, 0)  = 1"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                'ȡ���û��������
                strAdjustSQL = "Update " & gstrUserName & ".�ϻ���Ա�� b Set ϵͳ�������� = Null Where Nvl(ϵͳ��������, 0) = 1"
            End If
            '����SQL�Լ�����
            lngAll = rsObjectSQL.RecordCount + 2
            '����������û�
            If blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, "����������û�", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
            '����������û�
            Do While Not rsObjectSQL.EOF
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, IIf(blnLock, "�����û��˻�", "�����û��˻�"), (lngSort * 100 / lngAll), rsObjectSQL!SQL, 0)
                rsObjectSQL.MoveNext
            Loop
            'ȡ���û��������
            If Not blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, "ȡ���û��������", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
        Case FCT_��̨��ҵ
            If blnLock Then
                '���ͣ�õĺ�̨��ҵ
                strAdjustSQL = "Update Zlautojobs a" & vbNewLine & _
                                        "Set A.ϵͳ����ͣ�� = 1" & vbNewLine & _
                                        "Where Exists (Select 1 From All_Jobs b Where B.Broken = 'N' And B.Job = A.��ҵ��)"
                'ͣ�ú�̨��ҵ
                strSQL = "Select 'Dbms_Job.Broken(' || ��ҵ�� || ', True)' SQL" & vbNewLine & _
                                "From Zlautojobs a, All_Jobs b" & vbNewLine & _
                                "Where B.Job = A.��ҵ�� And B.Broken = 'N'"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                'ͣ��ϵͳ����
                strSQL = "Select Job_Name, 'Call dbms_scheduler.disable(' || Chr(39) || A.Owner|| '.'||Job_Name || Chr(39) || ')'  SQL" & vbNewLine & _
                            "From Dba_Scheduler_Jobs a" & vbNewLine & _
                            "Where A.Job_Name In ('GATHER_STATS_JOB', 'AUTO_SPACE_ADVISOR_JOB') And Enabled = 'TRUE'"
                Set rsSysScheduler = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                If rsSysScheduler.RecordCount <> 0 Then lngAll = 1: strTmp = ""
                lngAll = lngAll + rsObjectSQL.RecordCount + rsSysScheduler.RecordCount
            Else
                '���ú�̨��ҵ
                strSQL = "Select 'Dbms_Job.Broken(' || ��ҵ�� || ', False)' Sql From Zlautojobs Where ϵͳ����ͣ�� = 1"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                lngAll = rsObjectSQL.RecordCount
                'ȡ��ͣ�ñ��
                strAdjustSQL = "Update Zlautojobs Set ϵͳ����ͣ�� = Null Where Nvl(ϵͳ����ͣ��, 0) = 1"
                '����ϵͳ����
                strSQL = "select ���� from zlUpgradeConfig where ��Ŀ='���õ�ϵͳ����'"
                Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                strTmp = rsTmp!���� & ""
                If strTmp <> "" Then
                    lngAll = lngAll + 1
                    arrTmp = Split(strTmp, ",")
                    strTmp = ""
                    For i = LBound(arrTmp) To UBound(arrTmp)
                        strTmp = strTmp & ",'" & arrTmp(i) & "'"
                    Next
                    strTmp = Mid(strTmp, 2)
                    strSQL = "Select Job_Name, 'Call dbms_scheduler.enable(' || Chr(39) || A.Owner|| '.'||Job_Name || Chr(39) || ')' SQL" & vbNewLine & _
                            "From Dba_Scheduler_Jobs a" & vbNewLine & _
                            "Where A.Job_Name In (" & strTmp & ")"
                    Set rsSysScheduler = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                    lngAll = lngAll + rsSysScheduler.RecordCount
                End If
            End If
            '����SQL�Լ�����
            lngAll = lngAll + 1
            '���ͣ�õĺ�̨��ҵ
            If blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, "���ͣ�õĺ�̨��ҵ", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
            'ͣ�û����ú�̨��ҵ
            Do While Not rsObjectSQL.EOF
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, IIf(blnLock, "ͣ�ú�̨��ҵ", "���ú�̨��ҵ"), (lngSort * 100 / lngAll), rsObjectSQL!SQL, 0)
                rsObjectSQL.MoveNext
            Loop
            'ȡ��ͣ�ñ��
            If Not blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, "ȡ��ͣ�ñ��", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
            'ͣ�û�����ϵͳ����
            If Not rsSysScheduler Is Nothing Then
                Do While Not rsSysScheduler.EOF
                    lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, IIf(blnLock, "ͣ��ϵͳ����", "����ϵͳ����"), (lngSort * 100 / lngAll), rsSysScheduler!SQL, 1)
                    If blnLock Then
                        strTmp = strTmp & "," & rsSysScheduler!Job_Name
                    End If
                    rsSysScheduler.MoveNext
                Loop
            End If
            If blnLock And strTmp <> "" Then strTmp = Mid(strTmp, 2)
            '���ͣ��ϵͳ���Ȼ�ȡ��ϵͳ����ͣ�ñ��
            If strTmp <> "" Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, IIf(blnLock, "���ͣ��ϵͳ����", "ȡ��ϵͳ����ͣ�ñ��"), (lngSort * 100 / lngAll), "Update Zlupgradeconfig Set ���� = " & IIf(blnLock, "'" & strTmp & "'", "Null") & " Where ��Ŀ = '���õ�ϵͳ����'", 0)
            End If
        Case FCT_������
            If blnLock Then
                '��ǽ��õĴ�����
                strAdjustSQL = "Insert Into Zltriggers" & vbNewLine & _
                                        "  (������, ����)" & vbNewLine & _
                                        "  Select Owner, Trigger_Name" & vbNewLine & _
                                        "  From Dba_Triggers a" & vbNewLine & _
                                        "  Where A.Owner In (Select Upper(������)" & vbNewLine & _
                                        "                    From Zlbakspaces" & vbNewLine & _
                                        "                    Union" & vbNewLine & _
                                        "                    Select Upper(������)" & vbNewLine & _
                                        "                    From Zlsystems" & vbNewLine & _
                                        "                    Union" & vbNewLine & _
                                        "                    Select 'ZLTOOLS' From Dual) And A.Status = 'ENABLED'"
                '���ô�����
                strSQL = "Select 'alter trigger ' || Owner || '.' || Trigger_Name || ' disable' SQL" & vbNewLine & _
                                "From Dba_Triggers a" & vbNewLine & _
                                "Where A.Owner In (Select Upper(������)" & vbNewLine & _
                                "                  From Zlbakspaces" & vbNewLine & _
                                "                  Union" & vbNewLine & _
                                "                  Select Upper(������)" & vbNewLine & _
                                "                  From Zlsystems" & vbNewLine & _
                                "                  Union" & vbNewLine & _
                                "                  Select 'ZLTOOLS' From Dual) And A.Status = 'ENABLED'"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
            Else
                '���ô�����
                strSQL = "Select 'alter trigger ' || ������ || '.' || ���� || ' disable'  SQL From Zltriggers"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                'ɾ�����������ñ�־
                strAdjustSQL = "Delete ZLTOOLS.ZLTriggers"
            End If
            '����SQL�Լ�����
            lngAll = rsObjectSQL.RecordCount + 2
            '��ǽ��õĴ�����
            If blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, "��ǽ��õĴ�����", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
            '���û���ô�����
            Do While Not rsObjectSQL.EOF
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, IIf(blnLock, "ͣ�ú�̨��ҵ", "���ú�̨��ҵ"), (lngSort * 100 / lngAll), rsObjectSQL!SQL, 1)
                rsObjectSQL.MoveNext
            Loop
            'ɾ�����������ñ�־
            If Not blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, "ɾ�����������ñ�־", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
    End Select
    '��¼ִ�����
    lngSort = lngSort + 1: rsSQLStep.AddNew Array("���", "����", "����", "����", "DBA"), Array(lngSort, "ִ�б��", (lngSort * 100 / lngAll), "Update Zlupgradeconfig Set ���� = " & IIf(blnLock, 0, 1) & " Where ��Ŀ = '" & strInfoName & "'", 0)
    Set FuncControlSQL = rsSQLStep
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub SetProgressState(ByVal lngValue As Long, Optional ByVal strStep As String)
'���ܣ����ý�����
'������lngValue=���ȣ�-1�����ؽ�����   >=0����ʾ����������չʾ����
'          strStep=��ǰ������
    If lngValue = -1 Then
        pgbState.Visible = False: lblState.Visible = False: pgbState.value = 0
    Else
        If lngValue > 100 Then
            pgbState.Visible = False
            lblState.Caption = "������ϣ�"
            Sleep (1000)
            lblState.Visible = False
        Else
            pgbState.value = lngValue
            If lblState.Caption <> strStep Then
                lblState.Caption = strStep
                Call SetCtrlPosOnLine(False, 0, lblState, 0, pgbState)
            End If
            pgbState.Visible = True: lblState.Visible = True
        End If
    End If
End Sub

Private Function VilidateUpgrade(Optional ByVal lngRow As Long) As Boolean
    Dim i As Long, strMaxVer As String, strCurMaxVer As String
    Dim strMaxTools As String
    Dim blnUp As Boolean
    Dim strFilter As String, strFilterSys As String
    Dim lngBegin  As Long, lngEnd As Long
    
    On Error GoTo errH
    pgbGetUpInfo.value = 0: pgbGetUpInfo.Visible = True
    With vsSysSel
        If lngRow > .FixedRows Then
            strFilterSys = "ϵͳ���=" & Val(.TextMatrix(i, Col_���))
            lngBegin = lngRow: lngEnd = lngRow
        Else '�����߶�ȡʱ�������е�ϵͳ����ˢ��
            lngBegin = .FixedRows: lngEnd = .Rows - 1
        End If
        If lngRow <= .FixedRows Or mrsSysUpFiles Is Nothing Then '��ȡ��Ǩ�ļ�,��ʼ��
            Set mrsSysUpFiles = GetUpgradeFiles(Nothing, -1, "", "")
        Else '��յ�ǰϵͳ������
            Call RecDelete(mrsSysUpFiles, strFilterSys)
        End If
        '����ϴ���Ǩ�����Ϣ
        '"��Ǩ���", "��ֹ��Ϣ","��ǰ��Ǩ���", "��ǰ��ֹ��Ϣ"�����
        Call RecUpdate(mrsSysInfo, strFilterSys, "����", 1, "Ŀ��汾", Null, "���հ汾", Null, "������", 1, "�����", Null, "����ǰ����", 1, "��ǰ�����", Null)
        '�ϴγ�����Ǩδ������ɵĲ�����ǰִ��
        Call RecUpdate(mrsSysInfo, strFilterSys & IIf(strFilterSys <> "", " And ", "") & "��Ǩ���=1", "����ǰ����", 0, "��ǰ�����", "ϵͳ���һ�������������ɹ������ܽ�����ǰִ�У�")
        .Cell(flexcpText, lngBegin, Col_Ŀ��汾, lngEnd, Col_�����) = ""
        .Cell(flexcpForeColor, lngBegin, Col_Ŀ��汾, lngEnd, Col_�����) = &H80000008
        'ǰ��׼��
        pgbGetUpInfo.value = 5
        For i = lngBegin To lngEnd
            mrsSysInfo.Filter = "ϵͳ���=" & Val(.TextMatrix(i, Col_���))
            mrsSysInfo.Update "�����ļ�", .Cell(flexcpData, i, Col_�����ļ�)
            strMaxVer = "": strCurMaxVer = ""
            If .Cell(flexcpData, i, Col_�����ļ�) <> "" Then
                Set mrsSysUpFiles = GetUpgradeFiles(mrsSysUpFiles, Val(.TextMatrix(i, Col_���)), .TextMatrix(i, Col_��ǰ�汾), mrsSysInfo!�����ļ�, mrsSysInfo!��ֹ��Ϣ & "", mrsSysInfo!��ǰ��ֹ��Ϣ & "", strMaxVer, strCurMaxVer)
            End If
            mrsSysInfo.Update Array("���հ汾", "Ŀ��汾"), Array(strMaxVer, strCurMaxVer)
            pgbGetUpInfo.value = (i - lngBegin + 1) / (lngEnd - lngBegin + 1) * 75 + 5
        Next
    End With
    pgbGetUpInfo.value = 80
    Call RecToLog(mrsSysUpFiles, "ϵͳ���,FullSPVer,SysType,FileType", "�ļ���¼��")
    mrsSysInfo.Filter = "ϵͳ���=0"
    strMaxTools = IIf(mrsSysInfo!Ŀ��汾 & "" = "", mrsSysInfo!ϵͳ�汾��, mrsSysInfo!Ŀ��汾)
    mrsSysInfo.Filter = strFilterSys & IIf(strFilterSys <> "", " And ", "") & "������=1"
    Do While Not mrsSysInfo.EOF
        If mrsSysInfo!ϵͳ��� <> 0 Then
            If mrsSysInfo!Ŀ��汾 & "" <> "" Then
                '������Ŀ��汾֧�ֲ���Ӧ��ϵͳ��Ǩ��Ŀ��汾
                If VerFull(mrsSysInfo!Ŀ��汾) <= VerFull("10.34.0") Then '10.34.0�Ժ󣬹�������Ӧ��ϵͳ�汾����Ӧ��ֻ�Ƚ����汾
                    If VerFull(GetPrimaryVer(GetContractVersion(mrsSysInfo!Ŀ��汾))) > VerFull(GetPrimaryVer(strMaxTools)) Then
                        mrsSysInfo.Update Array("������", "�����"), Array(0, "�����߲���֧��Ӧ��ϵͳ��Ǩ��""" & mrsSysInfo!Ŀ��汾 & """!")
                    End If
                Else '10.34.0�Ժ󣬹�������Ӧ��ϵͳ�汾�Ѿ���Ӧ�����ԱȽ�SP�汾
                    If VerFull(GetPrimaryVer(mrsSysInfo!Ŀ��汾)) > VerFull(GetPrimaryVer(strMaxTools)) Then
                        mrsSysInfo.Update Array("������", "�����"), Array(0, "�����߲���֧��Ӧ��ϵͳ��Ǩ��""" & mrsSysInfo!Ŀ��汾 & """!")
                    End If
                End If
            Else
                mrsSysInfo.Update Array("������", "�����"), Array(0, "��ϵͳû�п�ִ�е������ű���������Ǩ!")
            End If
        ElseIf mrsSysInfo!Ŀ��汾 & "" = "" Then
            mrsSysInfo.Update Array("������", "�����"), Array(0, "������û�п�ִ�е������ű���������Ǩ!")
        End If
        mrsSysInfo.MoveNext
    Loop
    pgbGetUpInfo.value = 85
    '���ж�Ӧ��ϵͳ�ܷ񳣹���Ǩ��Ӧ�ò��ܳ�����Ǩ��������ǰ��Ǩ
    Call RecUpdate(mrsSysInfo, strFilterSys & IIf(strFilterSys <> "", " And ", "") & "������=0", "����ǰ����", 0)
    Call RecUpdate(mrsSysInfo, strFilterSys & IIf(strFilterSys <> "", " And ", "") & "����ǰ����=0 And ��ǰ�����=Null", "��ǰ�����", "!�����")
    pgbGetUpInfo.value = 90
    '��ȡ��ǰִ�е�Ŀ��汾
    If optUpMode(1).value Then
        mrsSysInfo.Filter = strFilterSys & IIf(strFilterSys <> "", " And ", "") & "����ǰ����=1"
        Do While Not mrsSysInfo.EOF
            strFilter = "ϵͳ���=" & mrsSysInfo!ϵͳ��� & " And SysType<>" & ST_History & " And FileType=" & FT_Before
            mrsSysUpFiles.Filter = strFilter: mrsSysUpFiles.Sort = "FullSPVer Desc": strMaxVer = ""
            If Not mrsSysUpFiles.EOF Then
                strMaxVer = mrsSysUpFiles!SPVer
                mrsSysUpFiles.Filter = strFilter & " And ���ð汾>'" & VerFull(mrsSysInfo!ϵͳ�汾��) & "'": mrsSysUpFiles.Sort = "FullSPVer"
                If Not mrsSysUpFiles.EOF Then
                    mrsSysUpFiles.Filter = strFilter & " And FullSPVer<'" & mrsSysUpFiles!FullSPVer & "'": mrsSysUpFiles.Sort = "FullSPVer Desc"
                    If Not mrsSysUpFiles.EOF Then
                        strMaxVer = mrsSysUpFiles!SPVer
                    Else
                        strMaxVer = ""
                        mrsSysInfo.Update Array("����ǰ����", "��ǰ�����"), Array(0, "û�п�ִ�е���ǰ�����ű���������ǰ��Ǩ��")
                    End If
                End If
            Else
                mrsSysInfo.Update Array("����ǰ����", "��ǰ�����"), Array(0, "û����ǰ�����ű���������ǰ��Ǩ��")
            End If
            mrsSysInfo.Update "��ǰĿ��汾", strMaxVer
            'ɾ������ǰ�ű�,��ɾ����ʷ����Ҫ����Ϊ��ʷ����ܰ汾�ϵͣ���Ҫ�����ȡ����ʱ��Ҫ�����Ľű�����ȡ�ϴη�����ֹ�Ժ�Ľű�
            Call RecDelete(mrsSysUpFiles, "ϵͳ���=" & mrsSysInfo!ϵͳ��� & " And SysType<>" & ST_History & " And FileType<>" & FT_Before)
            'ɾ��������ǰĿ��汾����ǰ�����ű�
            Call RecDelete(mrsSysUpFiles, strFilter & " And FullSPVer>'" & VerFull(strMaxVer) & "'")
            mrsSysInfo.MoveNext
        Loop
    End If
    '����չ��
    pgbGetUpInfo.value = 95
    With vsSysSel
        For i = lngBegin To lngEnd
            mrsSysInfo.Filter = "ϵͳ���=" & Val(.TextMatrix(i, Col_���))
            .RowData(i) = Val(IIf(optUpMode(1).value, mrsSysInfo!����ǰ����, mrsSysInfo!������) & "")
            If .RowData(i) = 0 And Val(.TextMatrix(i, Col_Sel)) <> 0 Then
                .TextMatrix(i, Col_Sel) = 0
            ElseIf Val(.TextMatrix(i, Col_Sel)) <> 0 Then
                blnUp = True
            End If
            mrsSysInfo.Update "����", IIf(Val(.TextMatrix(i, Col_Sel)) <> 0, 1, 0)
            .TextMatrix(i, Col_Ŀ��汾) = VerPAD(IIf(optUpMode(1).value, mrsSysInfo!��ǰĿ��汾, mrsSysInfo!Ŀ��汾) & "")
            .TextMatrix(i, Col_�����) = IIf(optUpMode(1).value, mrsSysInfo!��ǰ�����, mrsSysInfo!�����) & ""
            If .RowData(i) = 0 Then
                .Cell(flexcpForeColor, i, Col_Ŀ��汾, i, Col_�����) = &H2222B2 '��ש��
            End If
            If Val(.TextMatrix(i, Col_Sel)) = 0 Then '������������ɾ���ű�
                Call RecDelete(mrsSysUpFiles, "ϵͳ���=" & Val(vsSysSel.TextMatrix(i, Col_���)))
            End If
        Next
    End With
    pgbGetUpInfo.value = 100
    pgbGetUpInfo.Visible = False: pgbGetUpInfo.value = 0
    VilidateUpgrade = blnUp
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    pgbGetUpInfo.Visible = False
End Function

Private Sub SetSelBeach(Optional ByVal blnSel As Boolean)
'���ܣ���������ѡ��
'������blnSel=True������ѡ��False:����ȡ��
    Dim intSel As Integer
    Dim i As Long
    
    intSel = IIf(blnSel, 1, 0)
    With vsSysSel
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                .TextMatrix(i, Col_Sel) = intSel
            Else
                .TextMatrix(i, Col_Sel) = 0
            End If
        Next
    End With
End Sub
