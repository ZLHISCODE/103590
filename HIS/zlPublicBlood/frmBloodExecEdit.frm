VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBloodExecEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ѫִ�еǼ�"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14520
   Icon            =   "frmBloodExecEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPrompt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1440
      ScaleHeight     =   285
      ScaleWidth      =   9585
      TabIndex        =   28
      Top             =   8025
      Width           =   9585
      Begin VB.Label lblPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   30
         TabIndex        =   29
         Top             =   45
         Width           =   10500
      End
   End
   Begin VB.PictureBox picinfo 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   7020
      ScaleHeight     =   180
      ScaleWidth      =   3615
      TabIndex        =   26
      Top             =   90
      Width           =   3615
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "���ѣ���ע��ʼ������4h�����ѪҺ��ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   3525
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   7380
      Left            =   60
      ScaleHeight     =   7380
      ScaleWidth      =   14415
      TabIndex        =   0
      Top             =   480
      Width           =   14415
      Begin VB.CheckBox chkSign 
         Caption         =   "����ͬʱ���ǩ������"
         Height          =   180
         Left            =   11610
         TabIndex        =   35
         Top             =   1545
         Width           =   2115
      End
      Begin VB.TextBox txtִ��ժҪ 
         Height          =   1215
         Left            =   180
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   6090
         Width           =   14130
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   825
         TabIndex        =   11
         Top             =   90
         Width           =   13485
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   825
         TabIndex        =   10
         Top             =   1620
         Width           =   13485
      End
      Begin VB.Frame fraExe 
         BorderStyle     =   0  'None
         Caption         =   "��ע�˶�"
         Height          =   1245
         Index           =   2
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   14340
         Begin VB.Timer TimeFlash 
            Interval        =   250
            Left            =   10335
            Top             =   180
         End
         Begin VB.TextBox txtCheck 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   2
            Left            =   7860
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "2012-11-21 10:20"
            Top             =   135
            Width           =   1815
         End
         Begin VB.TextBox txtCheck 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   1
            Left            =   4455
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "����Ա"
            Top             =   105
            Width           =   1800
         End
         Begin VB.TextBox txtCheck 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   0
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "����Ա"
            Top             =   105
            Width           =   1815
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
            Height          =   600
            Left            =   180
            TabIndex        =   3
            Top             =   525
            Width           =   14130
            _cx             =   24924
            _cy             =   1058
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
            BackColorSel    =   16761024
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   9
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmBloodExecEdit.frx":000C
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
         Begin VB.Image imgMore 
            Height          =   225
            Left            =   9705
            Picture         =   "frmBloodExecEdit.frx":0192
            Top             =   165
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Line linB 
            Index           =   2
            X1              =   7860
            X2              =   9675
            Y1              =   375
            Y2              =   375
         End
         Begin VB.Label lblExeTime 
            AutoSize        =   -1  'True
            Caption         =   "�˶�ʱ��"
            Height          =   180
            Index           =   2
            Left            =   7065
            TabIndex        =   9
            Top             =   150
            Width           =   720
         End
         Begin VB.Line linB 
            Index           =   1
            X1              =   4410
            X2              =   6225
            Y1              =   345
            Y2              =   345
         End
         Begin VB.Line linB 
            Index           =   0
            X1              =   960
            X2              =   2775
            Y1              =   345
            Y2              =   345
         End
         Begin VB.Label lblCheck 
            AutoSize        =   -1  'True
            Caption         =   "�� �� ��"
            Height          =   180
            Index           =   1
            Left            =   3660
            TabIndex        =   8
            Top             =   135
            Width           =   720
         End
         Begin VB.Label lblCheck 
            AutoSize        =   -1  'True
            Caption         =   "�� �� ��"
            Height          =   180
            Index           =   0
            Left            =   165
            TabIndex        =   7
            Top             =   135
            Width           =   720
         End
      End
      Begin VB.Frame Frame3 
         Height          =   45
         Left            =   825
         TabIndex        =   1
         Top             =   5880
         Width           =   13485
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfExec 
         Height          =   3705
         Left            =   180
         TabIndex        =   13
         Top             =   1860
         Width           =   14130
         _cx             =   24924
         _cy             =   6535
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
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBloodExecEdit.frx":0593
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
         OwnerDraw       =   4
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
         Begin VB.PictureBox picDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1650
            ScaleHeight     =   270
            ScaleWidth      =   1725
            TabIndex        =   19
            Top             =   1515
            Visible         =   0   'False
            Width           =   1755
            Begin VB.CommandButton cmdDate 
               Height          =   270
               Left            =   1470
               Picture         =   "frmBloodExecEdit.frx":0735
               Style           =   1  'Graphical
               TabIndex        =   20
               TabStop         =   0   'False
               ToolTipText     =   "�༭(F4)"
               Top             =   0
               Width           =   270
            End
            Begin MSMask.MaskEdBox mskʱ�� 
               Height          =   300
               Left            =   0
               TabIndex        =   21
               Top             =   30
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   529
               _Version        =   393216
               BorderStyle     =   0
               MaxLength       =   16
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "yyyy-MM-dd hh:mm"
               Mask            =   "####-##-## ##:##"
               PromptChar      =   "_"
            End
         End
         Begin VB.PictureBox picText 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4170
            ScaleHeight     =   270
            ScaleWidth      =   750
            TabIndex        =   17
            Top             =   1530
            Visible         =   0   'False
            Width           =   780
            Begin VB.TextBox TxtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Width           =   570
            End
         End
         Begin VB.ListBox lstSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   570
            ItemData        =   "frmBloodExecEdit.frx":082B
            Left            =   6330
            List            =   "frmBloodExecEdit.frx":0835
            TabIndex        =   16
            Top             =   1410
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.PictureBox picCbo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4095
            ScaleHeight     =   270
            ScaleWidth      =   1440
            TabIndex        =   14
            Top             =   2070
            Visible         =   0   'False
            Width           =   1470
            Begin VB.ComboBox cboEdit 
               Height          =   300
               Left            =   -15
               TabIndex        =   15
               Text            =   "cboEdit"
               Top             =   -15
               Width           =   1500
            End
         End
      End
      Begin MSComCtl2.MonthView dtpDate 
         Height          =   2160
         Left            =   5580
         TabIndex        =   34
         Top             =   5175
         Visible         =   0   'False
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   3810
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   0
         StartOfWeek     =   266338305
         TitleBackColor  =   -2147483636
         TitleForeColor  =   -2147483634
         TrailingForeColor=   -2147483637
         CurrentDate     =   37904
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��Ѫ�˶�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   60
         TabIndex        =   24
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��ѪѲ��"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   60
         TabIndex        =   23
         Top             =   1545
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ִ��ժҪ"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   60
         TabIndex        =   22
         Top             =   5790
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   8040
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBloodExecEdit.frx":0841
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23178
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
   Begin VB.PictureBox picHide 
      Height          =   465
      Left            =   7695
      ScaleHeight     =   405
      ScaleWidth      =   1770
      TabIndex        =   30
      Top             =   4830
      Visible         =   0   'False
      Width           =   1830
      Begin VB.TextBox txt�������� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   165
         TabIndex        =   32
         Top             =   45
         Width           =   1005
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   270
         TabIndex        =   31
         Top             =   45
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker dtpҪ��ʱ�� 
         Height          =   300
         Left            =   0
         TabIndex        =   33
         Top             =   -15
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   266338307
         CurrentDate     =   38082
      End
   End
   Begin XtremeCommandBars.CommandBars cbsExec 
      Left            =   0
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBloodExecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnAcTive As Boolean
Private mstrȱʡ��Ѫ��Ӧ As String, mstr��Ѫ��Ӧ As String
Private mlngModul As Long
Private mlng�շ�ID As Long
Private mlngҽ��ID As Long, mlng���ID As Long
Private mlng���ͺ� As Long
Private mlng����ID As Long
Private mlngִ�п���ID As Long
Private mstrPrivs As String
Private mblnOk As Boolean
Private mstr����ʱ�� As String 'ѪҺ�Ľ���ʱ��
Private mintѪ���� As Integer, mint��ִ��Ѫ���� As Integer
Private mintTimerCount As Integer
Private mblnReturn As Boolean  'ִ���˿�������ƥ�����
Private mblnOnlyRead As Boolean '�Ƿ���ֻ��ģʽ
Private mblnShow As Boolean  '�Ƿ��ڱ༭״̬
Private mintType As Integer    '�༭��ʽ
Private mrsPersons As ADODB.Recordset  '��Ա��Ϣ
Private mrsItems As ADODB.Recordset  '����������Ŀ
Private mblnChange As Boolean
Private mlngNoEditor As Long '����Ĳ��ܱ༭��ʼ��
Private mstr��ʼʱ�� As String
Private mblnFinish As Boolean  'ִ�еǼ��Զ����ҽ��ִ��

Public Function ShowEdit(ByVal frmParent As Object, ByVal lngModul As Enum_Inside_Program, ByVal lngҽ��ID As Long, _
    ByVal lng���ͺ� As Long, ByVal lng����id As Long, ByVal lng�շ�ID As Long, ByVal lngִ�п���ID As Long, Optional ByVal strPrivs As String, Optional ByVal blnOnlyRead As Boolean, Optional blnFinish As Boolean) As Boolean
    mblnOk = False
    mblnFinish = False
    mlngModul = lngModul
    mlngҽ��ID = lngҽ��ID
    mlng���ͺ� = lng���ͺ�
    mlng����ID = lng����id
    mlng�շ�ID = lng�շ�ID
    mlngִ�п���ID = lngִ�п���ID
    mstrPrivs = strPrivs
    mblnOnlyRead = blnOnlyRead
    mblnShow = False
    On Error Resume Next
    Me.Show 1, frmParent
    If Err <> 0 Then Err.Clear
    blnFinish = mblnFinish
    ShowEdit = mblnOk
End Function

Public Sub ViewExecution(ByVal frmParent As Object, ByVal lng�շ�ID As Long)
    '���ܲ鿴��Ѫִ��
    mlng�շ�ID = lng�շ�ID
    mblnOnlyRead = True
    On Error Resume Next
    Me.Show 1, frmParent
    If Err <> 0 Then Err.Clear
End Sub

Private Sub cboEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picCbo_KeyDown(KeyCode, Shift)
End Sub

Private Sub cbsExec_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnOk As Boolean
    Dim strCheckOper As String, strCheckTime As String, strCheckResult As String
    Dim strSQL As String
    Dim intCol As Integer
    
    On Error GoTo ErrHand
    Select Case Control.id
        Case conMenu_Manage_ThingAudit '�˶�
            If txtCheck(2).Text <> "" Then
                MsgBox "�ô�ѪҺ�Ѿ��˶ԣ��������ٴκ˶ԣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            blnOk = frmUserCheck.ShowMe(Me, mlngModul, mlng����ID, mlng����ID, mstr����ʱ��, "", True, ִ�к˶�)
            If blnOk = True Then
                strCheckOper = frmUserCheck.SendAndTakeOper
                strCheckTime = frmUserCheck.SendTime
                strCheckResult = frmUserCheck.CheckResult

                 strSQL = "Zl_ѪҺִ�м�¼_Check(" & mlng�շ�ID & ",'" & Split(strCheckOper, "'")(0) & "','" & Split(strCheckOper, "'")(1) & "',To_Date('" & strCheckTime & "','YYYY-MM-DD HH24:MI:SS'),'" & strCheckResult & "')"
                Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)

                txtCheck(0).Text = Split(strCheckOper, "'")(0)
                txtCheck(1).Text = Split(strCheckOper, "'")(1)
                txtCheck(2).Text = strCheckTime
                Call LoadCheckVsf
            End If
            mblnOk = blnOk
        Case conMenu_Manage_ThingDelAudit 'ȡ��
            strCheckOper = ""
            If txtCheck(0).Text <> UserInfo.���� And txtCheck(1).Text <> UserInfo.���� Then
                strCheckOper = gobjDatabase.UserIdentifyByUser(Me, "��ȡ���˶�ǰ�������������û�����������������֤��", 100, mlngModul, "ִ������Ǽ�", , True)
                If strCheckOper = "" Then Exit Sub
                If txtCheck(0).Text <> strCheckOper And txtCheck(1).Text <> strCheckOper Then
                    MsgBox "ֻ��ȡ���Լ��˶Ի򸴲��ѪҺ!", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                If MsgBox("��ȷ��Ҫȡ���˶���", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
            End If
            strSQL = "Zl_ѪҺִ�м�¼_Uncheck(" & mlng�շ�ID & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            txtCheck(0).Text = ""
            txtCheck(1).Text = ""
            txtCheck(2).Text = ""
            vsfCheck.Cell(flexcpText, vsfCheck.FixedRows, vsfCheck.FixedCols, vsfCheck.Rows - 1, vsfCheck.Cols - 1) = ""
            mblnOk = True
        Case conMenu_Edit_Clear
            If vsfExec.Row >= vsfExec.FixedRows Then
                If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ����")) <> "" Then Exit Sub
                blnOk = vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ִ��ʱ��")) <> ""
                Call HiddenEditControl
                For intCol = vsfExec.FixedCols To vsfExec.Cols - 1
                    If vsfExec.ColHidden(intCol) = False Then
                        vsfExec.TextMatrix(vsfExec.Row, intCol) = ""
                    End If
                Next
                If blnOk Then
                    mblnChange = True
                    Call ChangeDataState
                End If
            End If
        Case conMenu_Tool_Sign, conMenu_Tool_SignEarse 'ǩ������;ȡ��ǩ������
            Call HiddenEditControl
            Call SignData(Control.id = conMenu_Tool_Sign)
        Case conMenu_Edit_Transf_Save
            mblnOk = SaveData
            If mblnOk = True Then
                mblnShow = False
                '����ʱ�������ҽ���Ƿ��Ѿ����ִ�У���������Զ����ҽ��ִ��
                If AutoAdviceFinish = True Then
                    mblnFinish = True
                    Unload Me
                Else
                    mblnFinish = False
                    Call RefreshDate
                    mblnChange = False
                End If
            End If
        Case conMenu_Edit_Transf_Cancle
            mblnShow = False
            Call RefreshDate
            mblnChange = False
        Case conMenu_File_Exit
            mblnChange = False
            Unload Me
    End Select
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mblnAcTive = True Then Exit Sub
    Select Case Control.id
        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = mblnChange And Control.Visible
        Case conMenu_Edit_Clear
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = vsfExec.Row >= vsfExec.FixedRows
            If Control.Enabled = True Then
                Control.Enabled = Control.Visible And vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ����")) = ""
            End If
        Case conMenu_Manage_ThingAudit
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = txtCheck(2).Text = "" And Control.Visible
        Case conMenu_Manage_ThingDelAudit
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = txtCheck(2).Text <> "" And mstr��ʼʱ�� = ""
        Case conMenu_Tool_Sign 'ǩ������
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = vsfExec.Row >= vsfExec.FixedRows
            If Control.Enabled = True Then
                Control.Enabled = mblnChange = False And vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ����")) = "" And Control.Visible
            End If
        Case conMenu_Tool_SignEarse 'ȡ��ǩ������
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = vsfExec.Row >= vsfExec.FixedRows
            If Control.Enabled = True Then
                Control.Enabled = mblnChange = False And vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ����")) <> "" And Control.Visible
            End If
    End Select
End Sub

Private Sub cbsExec_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = Bottom + stbThis.Height
End Sub

Private Sub cbsExec_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsExec.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    On Error Resume Next
    With picBack
        .Left = lngScaleLeft + 30
        .Top = lngScaleTop + 60
        .Width = lngScaleRight - .Left
        .Height = lngScaleBottom - .Top
    End With
    
    With picPrompt
        .Top = Me.ScaleHeight - stbThis.Height + 60
        .Height = stbThis.Height - 120
        .Left = stbThis.Panels(2).Left + 60
        .Width = stbThis.Panels(2).Width - 120
    End With
    With lblPrompt
        .FontSize = Me.FontSize
        .Width = picPrompt.Width
        .Height = TextHeight("��")
        .Top = (picPrompt.Height - .Height) \ 2
    End With
End Sub

Private Sub cmdDate_Click()
    With dtpDate
        .Tag = "mskʱ��"
        .Left = picDate.Left + vsfExec.Left
        .Top = picDate.Top + picDate.Height + vsfExec.Top
        If IsDate(cmdDate.Tag) Then
            .Value = Format(cmdDate.Tag, "YYYY-MM-DD")
        Else
            .Value = Format(gobjDatabase.Currentdate, "YYYY-MM-DD")
        End If
        .Visible = True
        .ZOrder 0
    End With
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    If dtpDate.Tag = "mskʱ��" And mskʱ��.Visible = True Then
        If IsDate(mskʱ��.Text) Then
            strDate = Format(DateClicked, "YYYY-MM-DD") & " " & Mid(Format(mskʱ��.Text, "YYYY-MM-DD HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "YYYY-MM-DD") & " " & Mid(Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm"), 12, 5)
        End If
        mskʱ��.Text = Format(strDate, "YYYY-MM-DD HH:mm")
        dtpDate.Visible = False
        If picDate.Enabled And picDate.Visible Then picDate.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ���Ϊ���ݷָ�������¼�¼���ķָ�������˲�����¼��
    If KeyAscii = 39 Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call HiddenEditControl
    End If
End Sub

Private Sub HiddenEditControl()
    mintType = -1
    picDate.Visible = False
    picText.Visible = False
    lstSelect.Visible = False
    picCbo.Visible = False
    dtpDate.Visible = False
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim blnUnLoad As Boolean
    
    On Error GoTo ErrHand
    mstr��ʼʱ�� = ""
    mlngNoEditor = 0
    mintTimerCount = 0
    mblnAcTive = True
    mintType = -1
    mblnChange = False
    picinfo.Visible = Not mblnOnlyRead
    TimeFlash.Enabled = Not mblnOnlyRead
    txtִ��ժҪ.locked = mblnOnlyRead
    chkSign.Visible = Not mblnOnlyRead
    
    Call InitExecBar '�˵���ʼ��
    If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & Me.name & "\Form", "״̬") <> "" Then
        DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & Me.name & "\Form", "״̬"
    End If
    Call gobjComlib.RestoreWinState(Me, App.ProductName)
    strSQL = "Select ����ʱ��,����״̬,ִ�к˶���,ִ�к˶�ʱ��,ִ�и�����,ִ�и���ʱ�� From ѪҺ���ͼ�¼ where �շ�ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�շ�ID)
    If rsTmp.EOF Then
        MsgBox "ѪҺ��δ����,���ܽ���ִ�еǼǣ�", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    If gbln���պ����ִ�� = True And mblnOnlyRead = False Then
        If Not (Val("" & rsTmp!����״̬) = 1 Or Val("" & rsTmp!����״̬) = 3) Then
            MsgBox "ѪҺ��δ����,���ܽ���ִ�еǼǣ�", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    If IsDate("" & rsTmp!����ʱ��) Then
        mstr����ʱ�� = Format("" & rsTmp!����ʱ��, "YYYY-MM-DD HH:mm")
    Else
        mstr����ʱ�� = ""
    End If
    
    '����ѪҺ�˶���Ϣ
    txtCheck(0).Text = rsTmp!ִ�к˶��� & ""
    txtCheck(1).Text = rsTmp!ִ�и����� & ""
    txtCheck(2).Text = Format(rsTmp!ִ�к˶�ʱ�� & "", "YYYY-MM-DD HH:mm")
    Call LoadCheckVsf
    
     If Val(gobjDatabase.GetPara("����ִ�еǼ�ͬʱǩ��", 2200, 9005, "0")) = 0 Or mblnOnlyRead = True Then
        chkSign.Value = 0
     Else
        chkSign.Value = 1
     End If
    '��Ѫ��Ӧ��ȡ
    mstr��Ѫ��Ӧ = ""
    strSQL = "Select ����,ȱʡ��־ From ��Ѫ��Ӧ"
    Call gobjDatabase.OpenRecordset(rsTmp, strSQL, "��Ѫ��Ӧ")
    Do While Not rsTmp.EOF
        mstr��Ѫ��Ӧ = mstr��Ѫ��Ӧ & "'" & rsTmp!����
        If Val(rsTmp!ȱʡ��־) = 1 Then mstrȱʡ��Ѫ��Ӧ = "" & rsTmp!����
    rsTmp.MoveNext
    Loop
    If Left(mstr��Ѫ��Ӧ, 1) = "'" Then mstr��Ѫ��Ӧ = Mid(mstr��Ѫ��Ӧ, 2)
    '������Ϣ
    strSQL = "Select ID, ������, ����, ��λ, ��ֵ��, С�� From ����������Ŀ" & _
        " Where ����id = 7 And ������ In ('����', '����', '����ѹ', '����ѹ', '����')"
    Set mrsItems = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
    '��Ա��Ϣ
    Set mrsPersons = GetDataToPersons
    If RefreshDate(blnUnLoad) = False Then
        If blnUnLoad = True Then Unload Me: Exit Sub
    End If
    mblnAcTive = False
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function RefreshDate(Optional blnUnLoad As Boolean) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    Call HiddenEditControl
    '����ѪҺִ����Ϣ
    mstr��ʼʱ�� = ""
    Call LoadExecVsf
    If mblnOnlyRead = False Then
        '��ȡҽ��ִ����Ϣ
        If mstr��ʼʱ�� = "" Then 'δִ��
            '��ȡ�ϴ�ִ����Ϣ
            strSQL = _
                " Select Sum(��������) as curNum" & _
                " From ����ҽ��ִ��" & _
                " Where ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�)
            If Not rsTmp.EOF Then
                txt��������.Tag = Nvl(rsTmp!curNum, 0) 'ѪҺҽ����ִ�д����ܺͣ�ÿ��ִ��һ��Ѫ��ִ�д���Ϊ1
            End If
            
            '���㱾��ִ��Ӧ�õ�Ҫ��ʱ��
            strSQL = "Select A.��������,Nvl(B.���id, B.ID) ��ID, B.��ʼִ��ʱ��" & _
                " From ����ҽ������ A,����ҽ����¼ B" & _
                " Where A.ҽ��ID=B.ID And A.ҽ��ID=[1] And A.���ͺ�=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�)
            
            dtpҪ��ʱ��.Value = rsTmp!��ʼִ��ʱ��  '��Ѫҽ����Ϊһ����ִ�е�����
            txt��������.Text = Val(rsTmp!�������� & "")
            mlng���ID = rsTmp!��ID
            mintѪ���� = GetBloodNum
            mint��ִ��Ѫ���� = gobjComlib.FormatEx(Val(txt��������.Tag) * mintѪ���� / Val(txt��������.Text), 0) '�ϴ�ִ�е�Ѫ�������Ѿ���5λС��������������Ϳ�������
            If mint��ִ��Ѫ���� >= mintѪ���� Then
                MsgBox "��ҽ�����η�������ִ�� " & mintѪ���� & "������ǰ�Ѿ�ִ���� " & mint��ִ��Ѫ���� & " ����������ִ�С�", vbInformation, gstrSysName
                blnUnLoad = True: Exit Function
            End If
    
            txt��������.Text = 1 'ÿ��ִ��Ĭ��Ϊһ��
        Else '�Ѿ�ִ��
            '��ѯ��ִ�е�ѪҺ���ܴ���(���㱾��)
            strSQL = "Select " & _
                " Sum(��������) as curNum" & _
                " From ����ҽ��ִ��" & _
                " Where ִ��ʱ��<>[3] And ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�, CDate(mstr��ʼʱ��))
            If Not rsTmp.EOF Then
                txt��������.Tag = Nvl(rsTmp!curNum, 0) 'ʵ����ִ�е��������������㱾�Σ�
            End If
            
            strSQL = "Select A.Ҫ��ʱ��,Nvl(C.���id, C.ID) ��ID,A.ִ��ʱ��,A.��������,A.ִ��ժҪ,A.ִ�н��,A.ִ����,B.��������" & _
                " From ����ҽ��ִ�� A,����ҽ������ B,����ҽ����¼ C" & _
                " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And B.ҽ��ID=C.ID" & _
                " And A.ҽ��ID=[1] And A.���ͺ�=[2] And A.ִ��ʱ��=[3]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�, CDate(mstr��ʼʱ��))
            If rsTmp.EOF Then
                MsgBox "δ���ڲ���ҽ��ִ�����ҵ���ѪҺ��ִ�м�¼�����飡", vbInformation, gstrSysName
                blnUnLoad = True: Exit Function
            End If
            
            dtpҪ��ʱ��.Value = rsTmp!Ҫ��ʱ��
            txt��������.Text = gobjComlib.FormatEx(Nvl(rsTmp!��������), 5)
            txtִ��ժҪ.Text = "" & rsTmp!ִ��ժҪ
            txt��������.Text = Val(rsTmp!�������� & "")
            mlng���ID = rsTmp!��ID
            If Trim(vsfExec.TextMatrix(vsfExec.FixedRows, vsfExec.ColIndex("ִ����"))) = "" Then vsfExec.TextMatrix(vsfExec.FixedRows, vsfExec.ColIndex("ִ����")) = rsTmp!ִ���� & ""
            
            mintѪ���� = GetBloodNum
            mint��ִ��Ѫ���� = gobjComlib.FormatEx(Val(txt��������.Tag) * mintѪ���� / Val(txt��������.Text), 0) '���ε�ִ�е�Ѫ����
            txt��������.Text = gobjComlib.FormatEx(Val("" & rsTmp!��������) * mintѪ����, 0)
        End If
    Else
        strSQL = _
            " Select a.ִ��ժҪ" & vbNewLine & _
            " From ����ҽ��ִ�� a, ѪҺִ�м�¼ b" & vbNewLine & _
            " Where a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And a.ִ��ʱ�� = b.ִ��ʱ�� And b.��¼���� = 1 And Nvl(b.���, 0) = 0 And b.�շ�id = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�շ�ID)
        If Not rsTmp.EOF Then
            txtִ��ժҪ.Text = "" & rsTmp!ִ��ժҪ
        End If
    End If
    RefreshDate = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetBloodNum() As Integer
'��ȡ����ҽ�����͵�����
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHand
    strSQL = "Select Count(�շ�id)  ���� From ѪҺ���ͼ�¼ a, ѪҺ��Ѫ��¼ b Where a.�䷢id = b.Id And b.����id = [1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng���ID)
    GetBloodNum = rsTemp!����
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadCheckVsf()
'���ܣ�����ѪҺ�˶Ա����Ϣ
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim arrName, arrKey
    
    On Error GoTo ErrHand
    arrName = Array("ѪҺЧ��", "ѪҺ����", "��Ѫװ��", "����", "סԺ��", "����", "����", "Ѫ��", "Ѫ����", "ѪҺ����", "ѪҺ����")
    arrKey = Array("ѪҺЧ���Ƿ���Ч", "ѪҺ�����Ƿ����", "��Ѫװ���Ƿ����", "�����Ƿ�һ��", "סԺ���Ƿ�һ��", "�����Ƿ�һ��", "�����Ƿ�һ��", "Ѫ���Ƿ�һ��", "Ѫ�����Ƿ���ȷ", "ѪҺ�����Ƿ���ȷ", "ѪҺ�����Ƿ���ȷ")
    With vsfCheck
        .Rows = 2
        .Cols = 12
        .FixedCols = 1
        .FixedRows = 1
        .Redraw = flexRDNone
        .ColWidth(0) = 1500
        .TextMatrix(0, 0) = "�˲���(3��8��)"
        .TextMatrix(1, 0) = "�˲���"
        For i = 1 To .Cols - 1
            .TextMatrix(0, i) = CStr(arrName(i - 1))
            .ColKey(i) = CStr(arrKey(i - 1))
            .ColWidth(i) = 1125
        Next
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
        strSQL = " Select ѪҺЧ���Ƿ���Ч, ѪҺ�����Ƿ����, ��Ѫװ���Ƿ����, �����Ƿ�һ��, סԺ���Ƿ�һ��, �����Ƿ�һ��, �����Ƿ�һ��, Ѫ���Ƿ�һ��, Ѫ�����Ƿ���ȷ, ѪҺ�����Ƿ���ȷ, ѪҺ�����Ƿ���ȷ" & vbNewLine & _
            " From ѪҺ�˶Խ��" & vbNewLine & _
            " Where �շ�id = [1] And ���� = [2]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�շ�ID, 3)
        If rsTemp.RecordCount > 0 Then
            For i = 0 To rsTemp.Fields.Count - 1
                vsfCheck.TextMatrix(1, vsfCheck.ColIndex(rsTemp.Fields(i).name)) = IIf(Val("" & rsTemp.Fields(i).Value) = 1, "��", "")
            Next
        End If
        
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(-1) = 255
        .Redraw = flexRDDirect
    End With
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadExecVsf()
'���ܣ�����ѪҺִ����Ϣ
    Dim i As Integer, intRow As Integer
    Dim arrName, arrKey, arrColWidth
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim int��� As Integer, blnNULL As Boolean, intAddRow As Integer
    Dim strValue As String
    
    On Error GoTo ErrHand
    '��ʼ�����(�༭��)
    arrName = Array("ִ��ʱ��", "ִ����", "����", "��Ѫ��λ������©", "�ܵ���ϴ", "ʹ��ҩ��", "��Ѫ��Ӧ", "��Ӧʱ��", "����", "����", "����", "Ѫѹ", "��¼����", "���", "�Ǽ���", "�Ǽ�ʱ��", "ǩ����", "ǩ��ʱ��", "״̬")
    arrKey = Array("ִ��ʱ��", "ִ����", "����", "������©", "�ܵ���ϴ", "ʹ��ҩ��", "��Ѫ��Ӧ", "��Ӧʱ��", "����", "����", "����", "Ѫѹ", "��¼����", "���", "�Ǽ���", "�Ǽ�ʱ��", "ǩ����", "ǩ��ʱ��", "״̬")
    arrColWidth = Array(1755, 900, 570, 900, 525, 525, 885, 1755, 600, 720, 720, 840, 0, 0, 900, 0, 900, 0, 0)
    With vsfExec
        .Clear
        .Cols = 21
        .Rows = 6
        .Redraw = flexRDNone
        .FixedRows = 1
        .FixedCols = 2
        .RowHeight(0) = 255
        .RowHeightMin = 255
        .MergeCells = flexMergeFixedOnly
        .MergeCol(0) = True
        .MergeCol(1) = True
        .FocusRect = IIf(mblnOnlyRead = True, flexFocusNone, flexFocusSolid)
        .BackColorSel = vbBlue
        .HighLight = flexHighlightNever
        .SelectionMode = flexSelectionFree
        
        .TextMatrix(1, 0) = "��עǰ15����"
        .TextMatrix(2, 0) = "��ע����"
        .TextMatrix(3, 0) = "��ע����"
        .TextMatrix(4, 0) = "��ע����"
        .TextMatrix(5, 0) = "��ע����4Сʱ"
        
        .TextMatrix(1, 1) = "��עǰ15����"
        .TextMatrix(2, 1) = "15���Ӻ�"
        .TextMatrix(3, 1) = "1Сʱ"
        .TextMatrix(4, 1) = "��ע����"
        .TextMatrix(5, 1) = "��ע����4Сʱ"
        .ColWidth(0) = 810
        .ColWidth(1) = 690
        
        For i = 2 To .Cols - 1
            Select Case CStr(arrName(i - 2))
                Case "����"
                    .TextMatrix(0, i) = CStr(arrName(i - 2)) & vbLf & "(��)"
                Case "����"
                    .TextMatrix(0, i) = CStr(arrName(i - 2)) & vbLf & "(��/��)"
                Case "����"
                    .TextMatrix(0, i) = CStr(arrName(i - 2)) & vbLf & "(��/��)"
                Case "Ѫѹ"
                    .TextMatrix(0, i) = CStr(arrName(i - 2)) & vbLf & "(mmHg)"
                Case Else
                    .TextMatrix(0, i) = CStr(arrName(i - 2))
            End Select
            
            .ColKey(i) = CStr(arrKey(i - 2))
            .ColWidth(i) = Val(arrColWidth(i - 2))
        Next
        .ColHidden(.ColIndex("��¼����")) = True
        .ColHidden(.ColIndex("���")) = True
        .ColHidden(.ColIndex("�Ǽ���")) = False
        .ColHidden(.ColIndex("�Ǽ�ʱ��")) = True
        .ColHidden(.ColIndex("ǩ����")) = False
        .ColHidden(.ColIndex("ǩ��ʱ��")) = True
        .ColHidden(.ColIndex("״̬")) = True
        .FrozenCols = .ColIndex("ִ��ʱ��")
        .Cell(flexcpAlignment, 0, .FixedCols, 0, .Cols - 1) = flexAlignCenterCenter
        mlngNoEditor = .ColIndex("��¼����")
        .TextMatrix(.FixedRows, .ColIndex("��¼����")) = 1: .TextMatrix(.FixedRows, .ColIndex("���")) = 0: .TextMatrix(.FixedRows, .ColIndex("״̬")) = 0
        .TextMatrix(.FixedRows + 1, .ColIndex("��¼����")) = 2: .TextMatrix(.FixedRows + 1, .ColIndex("���")) = 0: .TextMatrix(.FixedRows + 1, .ColIndex("״̬")) = 0
        .TextMatrix(.FixedRows + 2, .ColIndex("��¼����")) = 2: .TextMatrix(.FixedRows + 2, .ColIndex("���")) = 1: .TextMatrix(.FixedRows + 2, .ColIndex("״̬")) = 0
        .TextMatrix(.FixedRows + 3, .ColIndex("��¼����")) = 3: .TextMatrix(.FixedRows + 3, .ColIndex("���")) = 0: .TextMatrix(.FixedRows + 3, .ColIndex("״̬")) = 0
        .TextMatrix(.FixedRows + 4, .ColIndex("��¼����")) = 4: .TextMatrix(.FixedRows + 4, .ColIndex("���")) = 0: .TextMatrix(.FixedRows + 4, .ColIndex("״̬")) = 0
        '���̶ܹ��е��и߲���ȷ��Ҫ�Զ�������
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        .Cell(flexcpFloodColor, .FixedRows, .FixedCols, .FixedRows, .Cols - 1) = vbBlack
        
        '��ȡ����
        strSQL = " Select ҽ��id, ���ͺ�, ��¼����, ���, ִ��ʱ��, ִ����, ִ�п���id, ����, ��Ѫ��Ӧ, ��Ӧʱ��, ��Ѫ��λ�Ƿ���© �Ƿ���©, ��Ѫ�ܵ���ϴ,�Ƿ�ʹ��ҩ��, ����, ����, ����, ����ѹ, ����ѹ, ժҪ, �Ǽ���," & vbNewLine & _
            "       �Ǽ�ʱ��, ǩ����, ǩ��ʱ��" & vbNewLine & _
            " From ѪҺִ�м�¼" & vbNewLine & _
            " Where �շ�id = [1] order by ��¼����,nvl(���,0)"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�շ�ID)
        Do While Not rsTmp.EOF
            Select Case Val("" & rsTmp!��¼����)
                Case 1
                    mstr��ʼʱ�� = Format("" & rsTmp!ִ��ʱ��, "YYYY-MM-DD HH:mm:ss")
                    intRow = .FixedRows
                Case 2
                    intRow = .FixedRows + 1 + Val("" & rsTmp!���)
                    If intRow > .Rows - 3 Then '��ȥ�����й̶���
                        intAddRow = (intRow - .Rows + 3)
                        .Rows = .Rows + intAddRow
                        For i = .Rows - intAddRow To .Rows - 1
                            .TextMatrix(i, .ColIndex("��¼����")) = 2
                            .TextMatrix(i, .ColIndex("���")) = Val("" & rsTmp!���) - (.Rows - i - 1)
                            .RowPosition(i) = .Rows - 3 - intAddRow + 1
                        Next
                    End If
                Case 3
                    intRow = .Rows - 2
                Case 4
                    intRow = .Rows - 1
            End Select
            .TextMatrix(intRow, .ColIndex("ִ��ʱ��")) = Format("" & rsTmp!ִ��ʱ��, "YYYY-MM-DD HH:mm")
            .TextMatrix(intRow, .ColIndex("ִ����")) = "" & rsTmp!ִ����
            Select Case Val("" & rsTmp!����)
                Case -1
                    .TextMatrix(intRow, .ColIndex("����")) = "����"
                Case -2
                    .TextMatrix(intRow, .ColIndex("����")) = "��ѹ"
                Case Else
                    .TextMatrix(intRow, .ColIndex("����")) = "" & rsTmp!����
            End Select
            arrName = Array("�Ƿ���©", "�Ƿ�ʹ��ҩ��", "��Ѫ�ܵ���ϴ")
            arrKey = Array("������©", "ʹ��ҩ��", "�ܵ���ϴ")
            For i = 0 To UBound(arrName)
                strValue = "" & rsTmp(CStr(arrName(i))).Value
                Select Case strValue
                Case "0"
                     .TextMatrix(intRow, .ColIndex(CStr(arrKey(i)))) = "��"
                Case "1"
                     .TextMatrix(intRow, .ColIndex(CStr(arrKey(i)))) = "��"
                Case Else
                    .TextMatrix(intRow, .ColIndex(CStr(arrKey(i)))) = ""
                End Select
            Next
            .TextMatrix(intRow, .ColIndex("��Ѫ��Ӧ")) = "" & rsTmp!��Ѫ��Ӧ
            .TextMatrix(intRow, .ColIndex("��Ӧʱ��")) = Format("" & rsTmp!��Ӧʱ��, "YYYY-MM-DD HH:mm")
            .TextMatrix(intRow, .ColIndex("����")) = "" & rsTmp!����
            .TextMatrix(intRow, .ColIndex("����")) = "" & rsTmp!����
            .TextMatrix(intRow, .ColIndex("����")) = "" & rsTmp!����
            .TextMatrix(intRow, .ColIndex("Ѫѹ")) = "" & rsTmp!����ѹ & "/" & rsTmp!����ѹ
            If .TextMatrix(intRow, .ColIndex("Ѫѹ")) = "/" Then .TextMatrix(intRow, .ColIndex("Ѫѹ")) = ""
            .TextMatrix(intRow, .ColIndex("��¼����")) = "" & rsTmp!��¼����
            .TextMatrix(intRow, .ColIndex("���")) = "" & rsTmp!���
            .TextMatrix(intRow, .ColIndex("�Ǽ���")) = "" & rsTmp!�Ǽ���
            .TextMatrix(intRow, .ColIndex("�Ǽ�ʱ��")) = Format("" & rsTmp!�Ǽ�ʱ��, "YYYY-MM-DD HH:mm:ss")
            .TextMatrix(intRow, .ColIndex("ǩ����")) = "" & rsTmp!ǩ����
            .TextMatrix(intRow, .ColIndex("ǩ��ʱ��")) = Format("" & rsTmp!ǩ��ʱ��, "YYYY-MM-DD HH:mm:ss")
            .TextMatrix(intRow, .ColIndex("״̬")) = 1
            '.Cell(flexcpForeColor, intRow, .FixedCols, intRow, .Cols - 1) = IIf(.TextMatrix(intRow, .ColIndex("ǩ����")) <> "", vbRed, vbBlack)
        rsTmp.MoveNext
        Loop
        '��ע������������ˣ���Ԥ��һ��
        blnNULL = False
        For intRow = .FixedRows + 1 To .Rows - 3
            int��� = Val(.TextMatrix(intRow, .ColIndex("���")))
            If Val(.TextMatrix(intRow, .ColIndex("״̬"))) = 0 Then
                blnNULL = True
                Exit For
            End If
        Next
        If blnNULL = False Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("��¼����")) = 2
            .TextMatrix(.Rows - 1, .ColIndex("���")) = int��� + 1
            .RowPosition(.Rows - 1) = .Rows - 3
        End If
        '���¸�ֵ��������
         For intRow = .FixedRows + 1 To .Rows - 3
            .TextMatrix(intRow, 0) = "��ע����"
            int��� = Val(.TextMatrix(intRow, .ColIndex("���")))
            If int��� = 0 Then
                .TextMatrix(intRow, 1) = "15���Ӻ�"
            Else
                .TextMatrix(intRow, 1) = int��� & "Сʱ"
            End If
         Next
         '���ǹ̶��е��и�����Ϊ��С�и�
        For i = .FixedRows To .Rows - 1
            .RowHeight(i) = 300
            .MergeRow(i) = True
        Next
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Redraw = flexRDDirect
        Call vsfExec_AfterRowColChange(0, 0, vsfExec.FixedRows, vsfExec.ColIndex("ִ��ʱ��"))
    End With
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
'���ܣ�����ִ�м�¼
    Dim intRow As Integer, intNewRow As Integer, intCol As Integer
    Dim str��ʼִ��ʱ�� As String, strִ��ʱ�� As String
    Dim dbl�������� As Double, dblʣ����� As Double
    Dim blnTrans As Boolean, strSQL As String
    Dim arrSQL, i As Integer
    Dim int���� As Integer, str����ѹ As String, str����ѹ As String
    Dim arrMsg As Variant
    Dim blnDelete As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strCurDate As String
    Dim blnUpFirst As Boolean
    
    On Error GoTo ErrHand
    If mintType <> -1 Then Call MoveNextCell(False, True)
    
    blnDelete = True
    With vsfExec
        For intRow = .FixedRows To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("ִ��ʱ��")) <> "" Then
                blnDelete = False
                Exit For
            End If
        Next
    End With
    
    If blnDelete = False Then
        'δִ�к˶Ա����Ⱥ˶�
        If txtCheck(2).Text = "" Then
            MsgBox "���Ƚ�����Ѫǰ�˶ԣ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        With vsfExec
            str��ʼִ��ʱ�� = .TextMatrix(.FixedRows, .ColIndex("ִ��ʱ��"))
            '��Ѫ��ʼִ��ʱ�䲻�ܿ�
            If str��ʼִ��ʱ�� = "" Then
                MsgBox "����д��עǰ15����ִ��ʱ��ʱ�䣡", vbInformation, gstrSysName
                .Row = .FixedRows: .Col = .ColIndex("ִ��ʱ��")
                .ShowCell .Row, .Col
                If .Enabled And .Visible Then .SetFocus
                Exit Function
            End If
            '��Ѫ��ʼʱ�䲻��С�ں˶�ʱ��
            If IsDate(str��ʼִ��ʱ��) And IsDate(txtCheck(2).Text) Then
                If CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm")) < CDate(Format(txtCheck(2).Text, "yyyy-MM-dd HH:mm")) Then
                    MsgBox "������ע��ʼִ��ʱ�䲻��С�ں˶�ʱ�䡣", vbInformation, gstrSysName
                    .Row = .FixedRows: .Col = .ColIndex("ִ��ʱ��")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
            End If
            '��Ѫʱ�䲻��С��ҽ��ִ��Ҫ��ʱ��
            If IsDate(str��ʼִ��ʱ��) And IsDate(dtpҪ��ʱ��.Value) Then
                If CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm")) < CDate(Format(dtpҪ��ʱ��.Value, "yyyy-MM-dd HH:mm")) Then
                    MsgBox "������ע��ʼִ��ʱ�䲻��С��ҽ��Ҫ��ִ��ʱ�� " & Format(dtpҪ��ʱ��.Value, "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .Row = .FixedRows: .Col = .ColIndex("ִ��ʱ��")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
            End If
            '117041:��ʼʱ����ͬ����ʱ��������һ��ʱ���Զ���һ��,�ұ��ο�ʼʱ�䲻��С����һ��ִ�п�ʼʱ��
            If IsDate(str��ʼִ��ʱ��) Then
                strSQL = "Select Max(ִ��ʱ��) Lastdate" & vbNewLine & _
                        "From ����ҽ��ִ��" & vbNewLine & _
                        "Where ҽ��id = [1] And ���ͺ� = [2] And ִ��ʱ�� Between [3] And [4]" & IIf(mstr��ʼʱ�� <> "", " And ִ��ʱ��<>[5]", "")
                If mstr��ʼʱ�� <> "" Then
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�, CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm")), CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm") & ":59"), CDate(mstr��ʼʱ��))
                Else
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�, CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm")), CDate(Format(str��ʼִ��ʱ��, "YYYY-MM-DD HH:mm") & ":59"))
                End If
                If IsDate(rsTmp!Lastdate & "") Then
                        str��ʼִ��ʱ�� = Format(DateAdd("s", 1, CDate(Format(rsTmp!Lastdate, "yyyy-MM-dd HH:mm:ss"))), "yyyy-MM-dd HH:mm:ss")
                End If
            End If
        
            For intRow = .FixedRows To .Rows - 1
                '��Ѫʱ���Ƿ�����ȷ�����ڸ�ʽ
                If IsDate(.TextMatrix(intRow, .ColIndex("ִ��ʱ��"))) = False And .TextMatrix(intRow, .ColIndex("ִ��ʱ��")) <> "" Then
                    MsgBox GetExecName(intRow) & "��ִ��ʱ�䲻����Ч�����ڸ�ʽ��", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("ִ��ʱ��")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                If Val(.TextMatrix(intRow, .ColIndex("��¼����"))) = 1 Then
                    strִ��ʱ�� = str��ʼִ��ʱ��
                Else
                    strִ��ʱ�� = .TextMatrix(intRow, .ColIndex("ִ��ʱ��"))
                End If
                
                '���¼������ע����ÿСʱ��¼�������¼����ע15���Ӻ�����
                If .TextMatrix(intRow, .ColIndex("ִ��ʱ��")) <> "" Then
                    If Val(.TextMatrix(intRow, .ColIndex("��¼����"))) = 2 And Val(.TextMatrix(intRow, .ColIndex("���"))) > 0 Then
                        If .TextMatrix(.FixedRows + 1, .ColIndex("ִ��ʱ��")) = "" Then
                            MsgBox "¼������ע15���Ӻ�ÿСʱѲ�Ӽ�¼�������¼����ע15���Ӻ��¼��", vbInformation, gstrSysName
                            .Row = .FixedRows + 1: .Col = .ColIndex("ִ��ʱ��")
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                        
                        '�����һСʱ�Ƿ�¼����Ѳ�Ӽ�¼
                        If Val(.TextMatrix(intRow - 1, .ColIndex("��¼����"))) = 2 Then
                            If .TextMatrix(intRow - 1, .ColIndex("ִ��ʱ��")) = "" Then
                                MsgBox "��ע15���Ӻ�ÿСʱѲ�Ӽ�¼������������¼����ע15���Ӻ�" & Val(.TextMatrix(intRow - 1, .ColIndex("���"))) & "СʱѲ�Ӽ�¼��", vbInformation, gstrSysName
                                .Row = intRow - 1: .Col = .ColIndex("ִ��ʱ��")
                                .ShowCell .Row, .Col
                                If .Enabled And .Visible Then .SetFocus
                                Exit Function
                            End If
                        End If
                    End If
                End If
                
                '��д����Ѫ��ʼʱ�䣬������дִ����
                If .TextMatrix(intRow, .ColIndex("ִ��ʱ��")) <> "" And .TextMatrix(intRow, .ColIndex("ִ����")) = "" Then
                    MsgBox GetExecName(intRow) & "��ִ���˲���Ϊ�գ�", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("ִ����")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                '���ٵ�У��
                If InStr(1, ",����,��ѹ,,", "," & .TextMatrix(intRow, .ColIndex("����")) & ",") = 0 Then
                    If LenB(StrConv(.TextMatrix(intRow, .ColIndex("����")), vbFromUnicode)) > 3 Or Not IsNumeric(.TextMatrix(intRow, .ColIndex("����"))) Then
                        MsgBox GetExecName(intRow) & "�ĵ���ֻ�������֣������ֻ����¼��3λ���֣�", vbInformation, gstrSysName
                        .Row = intRow: .Col = .ColIndex("����")
                        .ShowCell .Row, .Col
                        If .Enabled And .Visible Then .SetFocus
                        Exit Function
                    End If
                End If
                
                '���׶ε�ִ��ʱ�����С����һ�׶ε�ִ��ʱ��
                For intNewRow = intRow + 1 To .Rows - 1
                    If IsDate(.TextMatrix(intRow, .ColIndex("ִ��ʱ��"))) And IsDate(.TextMatrix(intNewRow, .ColIndex("ִ��ʱ��"))) Then
                        If CDate(Format(.TextMatrix(intRow, .ColIndex("ִ��ʱ��")), "YYYY-MM-DD HH:mm")) >= CDate(Format(.TextMatrix(intNewRow, .ColIndex("ִ��ʱ��")), "YYYY-MM-DD HH:mm")) Then
                            MsgBox GetExecName(intRow) & "��ִ��ʱ�����С��" & GetExecName(intNewRow) & "��ִ��ʱ�䣡", vbInformation, gstrSysName
                            .Row = intRow: .Col = .ColIndex("ִ��ʱ��")
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                    End If
                Next
                
                'δ��дִ��ʱ�䣬����¼����������Ŀ����Ҫ����д
                If .TextMatrix(intRow, .ColIndex("ִ��ʱ��")) = "" Then
                    For intCol = .FixedCols To mlngNoEditor - 1
                        If .ColHidden(intCol) = False And Trim(.TextMatrix(intRow, intCol)) <> "" And intCol <> .ColIndex("ִ��ʱ��") Then
                            MsgBox GetExecName(intRow) & "��ִ��ʱ��Ϊ�գ�����д��������Ŀ���ݣ�����дִ��ʱ�䣡", vbInformation, gstrSysName
                            .Row = intRow: .Col = .ColIndex("ִ��ʱ��")
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                    Next
                End If
                
                '����Ѫ��Ӧ�������¼����Ѫ��Ӧ�����ʱ��
                If .TextMatrix(intRow, .ColIndex("��Ѫ��Ӧ")) <> "" And .TextMatrix(intRow, .ColIndex("��Ѫ��Ӧ")) <> "��" Then
                    If .TextMatrix(intRow, .ColIndex("��Ӧʱ��")) = "" Then
                        MsgBox GetExecName(intRow) & "������Ѫ��Ӧʱ�������¼�����뷴Ӧʱ�䣡", vbInformation, gstrSysName
                        .Row = intRow: .Col = .ColIndex("��Ӧʱ��")
                        .ShowCell .Row, .Col
                        If .Enabled And .Visible Then .SetFocus
                        Exit Function
                    End If
                End If
                '��Ѫ��Ӧʱ��У��
                If .TextMatrix(intRow, .ColIndex("��Ӧʱ��")) <> "" Then
                    If IsDate(.TextMatrix(intRow, .ColIndex("��Ӧʱ��"))) = False Then
                        MsgBox GetExecName(intRow) & "����Ѫ��Ӧʱ�䲻����Ч�����ڸ�ʽ��", vbInformation, gstrSysName
                        .Row = intRow: .Col = .ColIndex("��Ӧʱ��")
                        .ShowCell .Row, .Col
                        If .Enabled And .Visible Then .SetFocus
                        Exit Function
                    End If
                    '��Ѫ��Ӧʱ�䲻��С�ڱ���ִ��ʱ��
                     If IsDate(strִ��ʱ��) Then
                        If CDate(Format(.TextMatrix(intRow, .ColIndex("��Ӧʱ��")), "YYYY-MM-DD HH:mm")) < CDate(Format(strִ��ʱ��, "YYYY-MM-DD HH:mm")) Then
                            MsgBox GetExecName(intRow) & "����Ѫ��Ӧʱ�䲻��С��ִ��ʱ�䣡", vbInformation, gstrSysName
                            .Row = intRow: .Col = .ColIndex("��Ӧʱ��")
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                     End If
                     '���׶ε���Ѫ��Ӧʱ�䲻�ܴ�����һ�׶ε���Ѫ��Ӧʱ��
                     For intNewRow = intRow + 1 To .Rows - 1
                        If IsDate(.TextMatrix(intNewRow, .ColIndex("ִ��ʱ��"))) Then
                            If CDate(Format(.TextMatrix(intRow, .ColIndex("��Ӧʱ��")), "YYYY-MM-DD HH:mm")) >= CDate(Format(.TextMatrix(intNewRow, .ColIndex("ִ��ʱ��")), "YYYY-MM-DD HH:mm")) Then
                                MsgBox GetExecName(intRow) & "����Ѫ��Ӧʱ�����С��" & GetExecName(intNewRow) & "��ִ��ʱ�䣡", vbInformation, gstrSysName
                                .Row = intRow: .Col = .ColIndex("��Ӧʱ��")
                                .ShowCell .Row, .Col
                                If .Enabled And .Visible Then .SetFocus
                                Exit Function
                            End If
                        End If
                     Next
                End If
                '���������������
                If .TextMatrix(intRow, .ColIndex("����")) <> "" And Not IsNumeric(.TextMatrix(intRow, .ColIndex("����"))) Then
                    MsgBox GetExecName(intRow) & "�����²�����Ч���ָ�ʽ��", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("����")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                If .TextMatrix(intRow, .ColIndex("����")) <> "" And Not IsNumeric(.TextMatrix(intRow, .ColIndex("����"))) Then
                    MsgBox GetExecName(intRow) & "������������Ч���ָ�ʽ��", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("����")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                If .TextMatrix(intRow, .ColIndex("����")) <> "" And Not IsNumeric(.TextMatrix(intRow, .ColIndex("����"))) Then
                    MsgBox GetExecName(intRow) & "�ĺ���������Ч���ָ�ʽ��", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("����")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                If .TextMatrix(intRow, .ColIndex("Ѫѹ")) <> "" And InStr(1, .TextMatrix(intRow, .ColIndex("Ѫѹ")), "/") = 0 Then
                    MsgBox GetExecName(intRow) & "��Ѫѹ������ЧѪѹ��ʽ��", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("Ѫѹ")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
            Next
            '¼���������Сʱ�������ʱ�䲻��Ϊ��
            If IsDate(.TextMatrix(.Rows - 1, .ColIndex("ִ��ʱ��"))) And .TextMatrix(.Rows - 2, .ColIndex("ִ��ʱ��")) = "" Then
                MsgBox "¼������ע������4Сʱ�������¼����Ѫ������", vbInformation, gstrSysName
                .Row = .Rows - 2: .Col = .ColIndex("ִ��ʱ��")
                .ShowCell .Row, .Col
                If .Enabled And .Visible Then .SetFocus
                Exit Function
            End If
        End With
        
        If gobjCommFun.ActualLen(txtִ��ժҪ.Text) > txtִ��ժҪ.MaxLength Then
            MsgBox "ִ��ժҪ���ݹ��࣬������� " & txtִ��ժҪ.MaxLength \ 2 & " �����ֻ� " & txtִ��ժҪ.MaxLength & " ���ַ���", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txtִ��ժҪ)
            Exit Function
        End If
        
        '����ִ�д�������
        dbl�������� = Val(txt��������.Text)
        dblʣ����� = gobjComlib.FormatEx(Val(txt��������.Text) - Val(txt��������.Tag), 5)
        If mintѪ���� > mint��ִ��Ѫ���� Then
            dbl�������� = gobjComlib.FormatEx(dblʣ����� / (mintѪ���� - mint��ִ��Ѫ����), 5)
        Else
            dbl�������� = gobjComlib.FormatEx(dbl�������� / mintѪ����, 5)
        End If
        If Val(txt��������.Tag) + dbl�������� > Val(txt��������.Text) Then
            dbl�������� = gobjComlib.FormatEx(Val(txt��������.Text) - Val(txt��������.Tag), 5)
        End If
        
        '��������
        Call SetMessages(arrMsg)
        
        blnUpFirst = False
        arrSQL = Array()
        If mstr��ʼʱ�� = "" Then
            strSQL = "ZL_����ҽ��ִ��_Insert(" & mlngҽ��ID & "," & mlng���ͺ� & "," & _
                "To_Date('" & Format(dtpҪ��ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                dbl�������� & ",'" & txtִ��ժҪ.Text & "','" & gobjCommFun.GetNeedName(vsfExec.TextMatrix(vsfExec.FixedRows, vsfExec.ColIndex("ִ����"))) & "'," & _
                "To_Date('" & Format(str��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                1 & "," & "0," & 1 & ",'','" & UserInfo.��� & "','" & UserInfo.���� & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        Else
            strSQL = "ZL_����ҽ��ִ��_Update(To_Date('" & mstr��ʼʱ�� & "','YYYY-MM-DD HH24:MI:SS')," & mlngҽ��ID & "," & mlng���ͺ� & "," & _
                "To_Date('" & Format(dtpҪ��ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                dbl�������� & ",'" & txtִ��ժҪ.Text & "','" & gobjCommFun.GetNeedName(vsfExec.TextMatrix(vsfExec.FixedRows, vsfExec.ColIndex("ִ����"))) & "'," & _
                "To_Date('" & Format(str��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & "," & 1 & ",NULL," & 1 & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
            blnUpFirst = Format(mstr��ʼʱ��, "yyyy-MM-dd HH:mm:ss") <> Format(str��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
        End If
        strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
        With vsfExec
            For intRow = .FixedRows To .Rows - 1
                If .TextMatrix(intRow, .ColIndex("ִ��ʱ��")) <> "" And (Val(.TextMatrix(intRow, .ColIndex("״̬"))) <> 1 Or (blnUpFirst = True And Val(.TextMatrix(intRow, .ColIndex("��¼����"))) = 1)) Then
                    Select Case .TextMatrix(intRow, .ColIndex("����"))
                        Case "��ѹ"
                            int���� = -2
                        Case "����"
                            int���� = -1
                        Case Else
                            int���� = Val(.TextMatrix(intRow, .ColIndex("����")))
                    End Select
                    If InStr(1, .TextMatrix(intRow, .ColIndex("Ѫѹ")), "/") <> 0 Then
                        str����ѹ = Mid(.TextMatrix(intRow, .ColIndex("Ѫѹ")), 1, InStr(1, .TextMatrix(intRow, .ColIndex("Ѫѹ")), "/") - 1)
                        str����ѹ = Mid(.TextMatrix(intRow, .ColIndex("Ѫѹ")), InStr(1, .TextMatrix(intRow, .ColIndex("Ѫѹ")), "/") + 1)
                    Else
                        str����ѹ = "": str����ѹ = ""
                    End If
                    If Val(.TextMatrix(intRow, .ColIndex("��¼����"))) = 1 Then
                        strִ��ʱ�� = str��ʼִ��ʱ��
                    Else
                        strִ��ʱ�� = Format(Format(.TextMatrix(intRow, .ColIndex("ִ��ʱ��")), "YYYY-MM-DD HH:mm") & ":" & Format(strCurDate, "ss"), "YYYY-MM-DD HH:mm:ss")
                    End If
                    strSQL = "zl_ѪҺִ�м�¼_Update(" & mlng�շ�ID & "," & Val(.TextMatrix(intRow, .ColIndex("��¼����"))) & "," & Val(.TextMatrix(intRow, .ColIndex("���"))) & ",To_Date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                        gobjCommFun.GetNeedName(.TextMatrix(intRow, .ColIndex("ִ����"))) & "'," & mlng����ID & "," & IIf(int���� = 0, "NULL", int����) & ",'" & .TextMatrix(intRow, .ColIndex("��Ѫ��Ӧ")) & "'," & _
                        IIf(.TextMatrix(intRow, .ColIndex("��Ӧʱ��")) = "", "NULL", "To_Date('" & .TextMatrix(intRow, .ColIndex("��Ӧʱ��")) & "','YYYY-MM-DD HH24:MI:SS')") & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("������©")) = "��", 0, IIf(.TextMatrix(intRow, .ColIndex("������©")) = "��", 1, "NULL")) & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("�ܵ���ϴ")) = "��", 0, IIf(.TextMatrix(intRow, .ColIndex("�ܵ���ϴ")) = "��", 1, "NULL")) & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("ʹ��ҩ��")) = "��", 0, IIf(.TextMatrix(intRow, .ColIndex("ʹ��ҩ��")) = "��", 1, "NULL")) & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("����")) = "", "NULL", .TextMatrix(intRow, .ColIndex("����"))) & "," & IIf(.TextMatrix(intRow, .ColIndex("����")) = "", "NULL", .TextMatrix(intRow, .ColIndex("����"))) & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("����")) = "", "NULL", .TextMatrix(intRow, .ColIndex("����"))) & "," & IIf(str����ѹ = "", "NULL", str����ѹ) & "," & _
                        IIf(str����ѹ = "", "NULL", str����ѹ) & ",'" & UserInfo.���� & "',NULL,'" & txtִ��ժҪ.Text & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    If chkSign.Value <> 0 Then '����ͬʱ�������ǩ������
                        strSQL = "Zl_ѪҺִ�м�¼_Sign(" & mlng�շ�ID & "," & Val(.TextMatrix(intRow, .ColIndex("��¼����"))) & "," & Val(.TextMatrix(intRow, .ColIndex("���"))) & ",'" & UserInfo.���� & "',1)"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    End If
                ElseIf InStr(1, ",0,4,", "," & Val(.TextMatrix(intRow, .ColIndex("״̬"))) & ",") = 0 And .TextMatrix(intRow, .ColIndex("ִ��ʱ��")) = "" Then
                    strSQL = "Zl_ѪҺִ�м�¼_Delete(" & mlng�շ�ID & "," & Val(.TextMatrix(intRow, .ColIndex("��¼����"))) & "," & Val(.TextMatrix(intRow, .ColIndex("���"))) & ")"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                End If
            Next
        End With
    Else
        blnDelete = False
        arrSQL = Array()
        With vsfExec
            For intRow = .FixedRows To .Rows - 1
                If InStr(1, ",0,4,", "," & Val(.TextMatrix(intRow, .ColIndex("״̬"))) & ",") = 0 Then
                    blnDelete = True
                    Exit For
                End If
            Next
            If blnDelete = True Then
                Call SetMessages(arrMsg, True)
                strSQL = "ZL_����ҽ��ִ��_Delete(" & mlngҽ��ID & "," & mlng���ͺ� & ",To_Date('" & mstr��ʼʱ�� & "','YYYY-MM-DD HH24:MI:SS'))"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                
                strSQL = "Zl_ѪҺִ�м�¼_Delete(" & mlng�շ�ID & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            Else
                GoTo GOEND
            End If
        End With
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then
            Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        End If
    Next
    '��Ϣ���ݱ���
    For i = 0 To UBound(arrMsg)
        If CStr(arrMsg(i)) <> "" Then
            Call gobjDatabase.ExecuteProcedure(CStr(arrMsg(i)), Me.Caption)
        End If
    Next
    gcnOracle.CommitTrans: blnTrans = False
GOEND:
    SaveData = True
    Exit Function
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetExecName(ByVal intRow As Integer) As String
'���ܣ���ȡִ�ж�Ӧ�Ľ׶�����
    Dim strName As String
    With vsfExec
        Select Case Val(.TextMatrix(intRow, .ColIndex("��¼����")))
            Case 1
                strName = "��עǰ15����"
            Case 2
                If Val(.TextMatrix(intRow, .ColIndex("���"))) <= 0 Then
                    strName = "��ע15���Ӻ�"
                Else
                    strName = "��ע15���Ӻ�" & Val(.TextMatrix(intRow, .ColIndex("���"))) & "Сʱ"
                End If
            Case 3
                strName = "��ע����"
            Case 4
                strName = "��ע������4Сʱ"
        End Select
    End With
    
    GetExecName = strName
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err.Clear
    On Error Resume Next
    Call gobjDatabase.SetPara("����ִ�еǼ�ͬʱǩ��", chkSign.Value, 2200, 9005)
    If Not mrsPersons Is Nothing Then
        If mrsPersons.State = adStateOpen Then mrsPersons.Close
        Set mrsPersons = Nothing
    End If
    If Not mrsItems Is Nothing Then
        If mrsItems.State = adStateOpen Then mrsItems.Close
        Set mrsItems = Nothing
    End If
    Call gobjComlib.SaveWinState(Me, App.ProductName)
    If Err <> 0 Then Err.Clear
End Sub

Private Sub lstSelect_DblClick()
    Call lstSelect_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lstSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    End If
End Sub

Private Sub mskʱ��_GotFocus()
    mskʱ��.SelStart = 0: mskʱ��.SelLength = Len(mskʱ��.Text)
End Sub

Private Sub mskʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picDate_KeyDown(KeyCode, Shift)
End Sub

Private Sub picCbo_GotFocus()
    If cboEdit.Enabled And cboEdit.Visible Then cboEdit.SetFocus
End Sub

Private Sub picCbo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell(, , True)
    End If
End Sub

Private Sub picDate_GotFocus()
    If mskʱ��.Visible And mskʱ��.Enabled Then mskʱ��.SetFocus
End Sub

Private Sub picDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    End If
End Sub

Private Sub picText_GotFocus()
    If TxtEdit.Enabled And TxtEdit.Visible Then TxtEdit.SetFocus
End Sub

Private Sub picText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    End If
End Sub

Private Sub TimeFlash_Timer()
    mintTimerCount = mintTimerCount + 1
    
    If mintTimerCount Mod 2 = 0 Then
        lblTitle.ForeColor = 0
    Else
        lblTitle.ForeColor = 255
    End If
    
    If mintTimerCount = 10 Then mintTimerCount = 0
End Sub

Private Sub TxtEdit_GotFocus()
    Call gobjControl.TxtSelAll(TxtEdit)
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picText_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtִ��ժҪ_Change()
    If mblnAcTive = True Then Exit Sub
    mblnChange = True
End Sub

Private Sub txtִ��ժҪ_GotFocus()
    Call gobjControl.TxtSelAll(txtִ��ժҪ)
End Sub

Private Sub vsfExec_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strֵ�� As String, intС�� As Integer, strTmp As String
    Dim arrName, i As Integer
    Dim blnMatch As Boolean
    
    Call ShowMsg("")
    On Error GoTo ErrHand
    If vsfExec.Cell(flexcpBackColor, NewRow, NewCol) <> 16772055 Then
        Call SetColBackColor(16772055)
    End If
    If mblnOnlyRead = True Then
        vsfExec.FocusRect = flexFocusNone
    Else
        If NewCol >= vsfExec.FixedCols And NewCol < mlngNoEditor Then
            vsfExec.FocusRect = flexFocusSolid
        Else
            vsfExec.FocusRect = flexFocusHeavy
        End If
    End If
    If vsfExec.TextMatrix(NewRow, vsfExec.ColIndex("ǩ����")) <> "" Then
        strTmp = "����ִ�м�¼�ѱ�����Ա[" & vsfExec.TextMatrix(NewRow, vsfExec.ColIndex("ǩ����")) & "]ǩ��"
        Call ShowMsg(strTmp, vbRed)
        Exit Sub
    End If
    Select Case NewCol
        Case vsfExec.ColIndex("����"), vsfExec.ColIndex("����"), vsfExec.ColIndex("����")
            strֵ�� = "": intС�� = -1
            mrsItems.Filter = "������='" & vsfExec.ColKey(NewCol) & "'"
            If Not mrsItems.EOF Then
                Select Case vsfExec.ColKey(NewCol)
                    Case "����"
                        blnMatch = mrsItems!��λ & "" = "��"
                    Case "����", "����"
                        blnMatch = mrsItems!��λ & "" = "��/��"
                    Case "����ѹ", "����ѹ"
                        blnMatch = mrsItems!��λ & "" = "mmHg"
                End Select
                If blnMatch = True Then
                    strֵ�� = mrsItems!��ֵ�� & ""
                    intС�� = Val(mrsItems!С�� & "")
                End If
            End If
            If InStr(1, strֵ��, ";") = 0 Or intС�� = -1 Then
             '�Ҳ�����ʹ��ȱʡֵ
                Select Case vsfExec.ColKey(NewCol)
                    Case "����"
                        If InStr(1, strֵ��, ";") = 0 Then strֵ�� = "35;42"
                        If intС�� = -1 Then intС�� = 1
                    Case "����"
                        If InStr(1, strֵ��, ";") = 0 Then strֵ�� = "20;300"
                        If intС�� = -1 Then intС�� = 0
                    Case "����"
                        If InStr(1, strֵ��, ";") = 0 Then strֵ�� = "15;50"
                        If intС�� = -1 Then intС�� = 0
                    Case "����ѹ", "����ѹ"
                        If InStr(1, strֵ��, ";") = 0 Then strֵ�� = "50;190"
                        If intС�� = -1 Then intС�� = 0:
                End Select
            End If
            
            If Left(strֵ��, 1) = "." Then strֵ�� = "0" & strֵ��
            strTmp = Replace(strֵ��, ";", " - ")
            If intС�� = 0 Then
                strTmp = "��¼�뷶ΧΪ " & strTmp & " ����"
            Else
                strTmp = "��¼�뷶ΧΪ " & strTmp & " ֮��������ɺ�" & intС�� & "λС��"
            End If
            Call ShowMsg(strTmp)
        Case vsfExec.ColIndex("Ѫѹ")
            arrName = Array("����ѹ", "����ѹ")
            For i = 0 To UBound(arrName)
                mrsItems.Filter = "������='" & vsfExec.ColKey(NewCol) & "'"
                If Not mrsItems.EOF Then
                    strֵ�� = mrsItems!��ֵ�� & ""
                    intС�� = Val(mrsItems!С�� & "")
                End If
                If InStr(1, strֵ��, ";") = 0 Then strֵ�� = "50;190"
                If intС�� = -1 Then intС�� = 0
                If Left(strֵ��, 1) = "." Then strֵ�� = "0" & strֵ��
                If i = 0 Then
                    strTmp = Replace(strֵ��, ";", " - ")
                Else
                    strTmp = strTmp & "/" & Replace(strֵ��, ";", " - ")
                End If
            Next
            strTmp = "��ʽΪ[����ѹ/����ѹ]����¼�뷶ΧΪ " & strTmp
            Call ShowMsg(strTmp)
    End Select
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfExec_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If mblnAcTive = True Then Exit Sub
    Call HiddenEditControl
End Sub

Private Sub vsfExec_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnAcTive = True Then Exit Sub
    If mintType = -1 Then Exit Sub
    Cancel = Not MoveNextCell(True, True)
End Sub

Private Sub vsfExec_DblClick()
    Call vsfExec_KeyDown(Asc("A"), 0)
End Sub

Private Sub vsfExec_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    If DrawTimeCell(hDC, Row, Col, Left, Top, Right, Bottom) = False Then Exit Sub
    Done = True
End Sub

Private Function DrawTimeCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Boolean
    Dim rc As RECT
    Dim rcUp As RECT
    Dim lngLoop As Long
    Dim lngSvrBkColor As Long
    Dim lngCenterTop As Long
    Dim lngPen As Long, lngOldPen As Long
    Dim lpPoint As POINTAPI
    If Not (Row = 0 And (Col = 0 Or Col = 1)) Then Exit Function
    With vsfExec
        rc.Left = Left
        rc.Top = Top
        rc.Bottom = Bottom - 1
        rc.Right = Right - Col
        lngSvrBkColor = &H8000000F
        Call SetBkColor(hDC, GetRBGFromOLEColor(lngSvrBkColor))
        Call ExtTextOut(hDC, rc.Left, rc.Top, 2, rc, " ", 1, lngLoop)
        
        '���1����ඥ�˺͵ڶ����Ҳ�Ͷ����ߣ��������м�Ľ����
        lngCenterTop = (.RowHeight(0) * .ColWidth(1)) / (.ColWidth(0) + .ColWidth(1)) \ 15 + 2
        '����
        lngPen = CreatePen(PS_SOLID, 1, vbBlack)
        lngOldPen = SelectObject(hDC, lngPen)
        '��ͼ
        If Col = 0 Then
            Call MoveToEx(hDC, rc.Left, rc.Top, lpPoint)
            Call LineTo(hDC, rc.Right, lngCenterTop)
            '9������ĸ߶ȺͿ�Ⱦ���180�����ؾ���12
            Call TextOut(hDC, rc.Left + 2, rc.Bottom - 12 - 2, "����", 4)
        Else
            '��ͼ
            Call MoveToEx(hDC, rc.Left, lngCenterTop, lpPoint)
            Call LineTo(hDC, rc.Right, rc.Bottom)
            '9������ĸ߶ȺͿ�Ⱦ���180�����ؾ���12
            Call TextOut(hDC, rc.Right - 24 - 2, rc.Top + 2, "����", 4)
        End If
        '��ԭ���ʲ�����
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
    End With
End Function

Private Sub vsfExec_EnterCell()
    
    If mblnAcTive = True Or mblnOnlyRead = True Then Exit Sub
    '���ر༭�Ŀؼ�
    picDate.Visible = False
    picText.Visible = False
    lstSelect.Visible = False
    picCbo.Visible = False
    dtpDate.Visible = False
    
    With vsfExec
        If mblnShow = True Then
            If .TextMatrix(.Row, .ColIndex("ǩ����")) <> "" Then Exit Sub
            'δ��дִ��ʱ�� ��������д�����
            If .TextMatrix(.Row, .ColIndex("ִ��ʱ��")) = "" And .Col <> .ColIndex("ִ��ʱ��") Then
                If .Col >= .FixedCols And .Col < mlngNoEditor Then
                    Call ShowMsg("��д��ִ��ʱ�䣬��������д������", vbRed)
                End If
                Exit Sub
            End If
            'Ҫ��д��Ѫ��Ӧʱ�䣬�������д��Ѫ��Ӧ
            If .Col = .ColIndex("��Ӧʱ��") And InStr(1, "'��'", "'" & Trim(.TextMatrix(.Row, .ColIndex("��Ѫ��Ӧ")))) <> 0 Then
                Exit Sub
            End If
            '��Ѫ�ܵ���ϴ��ֻ������Ѫǰ����Ѫ����д
            If Val(.TextMatrix(.Row, .ColIndex("��¼����"))) = 2 And .Col = .ColIndex("�ܵ���ϴ") Then Exit Sub
            '��Ѫ������4Сʱ������д ���١��Ƿ�ʹ��ҩ��Ƿ���©���ܵ���ϴ
            If Val(.TextMatrix(.Row, .ColIndex("��¼����"))) = 4 Then
                If .Col = .ColIndex("����") Or .Col = .ColIndex("������©") Or .Col = .ColIndex("ʹ��ҩ��") Or .Col = .ColIndex("�ܵ���ϴ") Then
                    Exit Sub
                End If
            End If
        End If
    End With
    If Not mblnShow Then Exit Sub
    '��ʼ��ʾ�ؼ�
    If vsfExec.Col < mlngNoEditor Then Call ShowInput
End Sub

Private Sub vsfExec_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (vsfCheck.Col >= vsfCheck.FixedRows And vsfExec.Row >= vsfExec.FixedRows) Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If mblnShow = False And vsfExec.Col = vsfExec.ColIndex("ִ��ʱ��") Then
            mblnShow = True
            Call vsfExec_EnterCell
        Else
            Call MoveNextCell
        End If
    ElseIf Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyDelete Or Shift <> 0) Then
            mblnShow = True
            Call vsfExec_EnterCell
    ElseIf KeyCode = vbKeyDelete Then
        If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ����")) <> "" Then Exit Sub
        If vsfExec.TextMatrix(vsfExec.Row, vsfExec.Col) <> "" And vsfExec.Col <> vsfExec.ColIndex("ִ��ʱ��") Then
            HiddenEditControl
            vsfExec.TextMatrix(vsfExec.Row, vsfExec.Col) = ""
            mblnChange = True
            Call ChangeDataState
        End If
    End If
End Sub

Public Function SetColBackColor(Optional ByVal lngColor As Long = 16772055) As Boolean
    '******************************************************************************************************************
    '����:�����еı���ɫ
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    On Error Resume Next
    
    For lngLoop = vsfExec.FixedCols To vsfExec.Cols - 1
        If vsfExec.ColHidden(lngLoop) = False Then
            vsfExec.Cell(flexcpBackColor, vsfExec.FixedRows, lngLoop, vsfExec.Rows - 1, lngLoop) = 16777215
        End If
    Next
    For lngLoop = vsfExec.FixedCols To vsfExec.Cols - 1
        If vsfExec.Cell(flexcpBackColor, vsfExec.Row, lngLoop, vsfExec.Row, lngLoop) = 16777215 Then
            vsfExec.Cell(flexcpBackColor, vsfExec.Row, lngLoop, vsfExec.Row, lngLoop) = lngColor
        End If
    Next
    If Err <> 0 Then Err.Clear
End Function

Private Sub ShowInput()
'��ʾ��Ӧ�ı༭�ؼ�
    Dim i As Integer, lngLegth As Long
    Dim strText As String
    Dim CellRect As RECT
    Dim lngFindCboIndex As Long
    Dim arrTmp
    
    With vsfExec
        If vsfExec.ColIsVisible(vsfExec.Col) = False Then
            vsfExec.LeftCol = vsfExec.Col
        End If
        If vsfExec.RowIsVisible(vsfExec.Row) = False Then
            vsfExec.TopRow = vsfExec.Row
        End If
        'ȷ���ؼ���λ��
        CellRect.Left = .CellLeft
        CellRect.Top = .CellTop
        CellRect.Right = .CellWidth - 10
        CellRect.Bottom = .CellHeight - 10
        strText = .TextMatrix(.Row, .Col)
        'ȷ��Ҫ��ʾ�Ŀؼ�
        Select Case .ColKey(.Col)
            Case "ִ��ʱ��", "��Ӧʱ��"
                mintType = 0
                picDate.Left = CellRect.Left
                picDate.Top = CellRect.Top
                picDate.Width = CellRect.Right
                picDate.Height = CellRect.Bottom
                picDate.BackColor = .BackColor
                picDate.BorderStyle = 0
                picDate.Visible = True
                picDate.ZOrder 0
                picDate.SetFocus
                picDate.Tag = .ColKey(.Col)
                '��ֵ
                If IsDate(strText) Then
                    mskʱ��.Text = Format(strText, "YYYY-MM-DD HH:mm")
                Else
                    mskʱ��.Text = "____-__-__ __:__"
                End If
                cmdDate.Tag = strText
            Case "������©", "ʹ��ҩ��", "��Ѫ��Ӧ", "�ܵ���ϴ"
                mintType = 1
                arrTmp = Split(mstr��Ѫ��Ӧ, "'")
                lstSelect.Clear
                If .ColKey(.Col) = "��Ѫ��Ӧ" And UBound(arrTmp) >= 0 Then
                    For i = 0 To UBound(arrTmp)
                        lstSelect.AddItem CStr(arrTmp(i))
                        If CStr(arrTmp(i)) = mstrȱʡ��Ѫ��Ӧ Then
                            lstSelect.Selected(lstSelect.NewIndex) = True
                        End If
                    Next
                Else
                    lstSelect.AddItem "��"
                    lstSelect.AddItem "��"
                End If
                lstSelect.Left = CellRect.Left
                lstSelect.Top = CellRect.Top + CellRect.Bottom
                lstSelect.Width = CellRect.Right
                lstSelect.Height = lstSelect.ListCount * (picText.TextHeight("��")) + picText.TextHeight("��") \ 3
                If lstSelect.Height < CellRect.Bottom Then lstSelect.Height = CellRect.Bottom
                
                '��ȡ��ĳ�������
                lngLegth = 0
                For i = 0 To lstSelect.ListCount - 1
                    If lngLegth < LenB(StrConv(lstSelect.List(i), vbFromUnicode)) Then
                        lngLegth = LenB(StrConv(lstSelect.List(i), vbFromUnicode))
                    End If
                    If strText = lstSelect.List(i) Then
                        lstSelect.Selected(i) = True
                    Else
                        lstSelect.Selected(i) = False
                    End If
                Next
                lstSelect.Width = lngLegth * picText.TextWidth("1") + 60    '������������Ϊ׼
                If lstSelect.Width < CellRect.Right Then lstSelect.Width = CellRect.Right
                If lstSelect.Height + lstSelect.Top > vsfExec.ClientHeight Then
                    If vsfExec.ClientHeight - CellRect.Top > vsfExec.ClientHeight - lstSelect.Top Then
                         lstSelect.Top = vsfExec.ClientHeight - lstSelect.Height
                         If lstSelect.Top < vsfExec.RowHeight(0) Then lstSelect.Top = vsfExec.RowHeight(0)
                         If lstSelect.Top + lstSelect.Height > vsfExec.ClientHeight Then
                             lstSelect.Height = vsfExec.ClientHeight - lstSelect.Top
                         End If
                    Else
                        lstSelect.Height = vsfExec.ClientHeight - lstSelect.Top
                    End If
                End If
                
                lstSelect.Visible = True
                lstSelect.ZOrder 0
                lstSelect.Tag = .ColKey(.Col)
                lstSelect.SetFocus
            Case "����", "����", "����", "Ѫѹ"
                mintType = 2
                picText.Left = CellRect.Left
                picText.Top = CellRect.Top
                picText.Width = CellRect.Right
                picText.Height = CellRect.Bottom
                picText.BackColor = .BackColor
                picText.BorderStyle = 0
                TxtEdit.Width = picText.Width
                TxtEdit.Height = picText.Height
                TxtEdit.Left = 0
                TxtEdit.Top = 0
                TxtEdit.Text = strText
                TxtEdit.Tag = strText
                picText.Visible = True
                picText.ZOrder 0
                picText.SetFocus
                picText.Tag = .ColKey(.Col)
            Case "����", "ִ����"
                mintType = 3
                cboEdit.Clear
                cboEdit.Text = ""
                cboEdit.locked = False
                cboEdit.Tag = strText
                If .ColKey(.Col) = "����" Then
                    cboEdit.AddItem 15: cboEdit.ItemData(cboEdit.NewIndex) = 15
                    cboEdit.AddItem 30: cboEdit.ItemData(cboEdit.NewIndex) = 30
                    cboEdit.AddItem "����": cboEdit.ItemData(cboEdit.NewIndex) = -1
                    cboEdit.AddItem "��ѹ": cboEdit.ItemData(cboEdit.NewIndex) = -2
                    gobjComlib.cbo.SetText cboEdit, strText
                Else
                    mrsPersons.Filter = ""
                    lngFindCboIndex = -1
                    Do While Not mrsPersons.EOF
                        cboEdit.AddItem mrsPersons!���� 'mrsPersons!��� & "-" & mrsPersons!����
                        cboEdit.ItemData(cboEdit.NewIndex) = Val("" & mrsPersons!id)
                        If strText = "" Then
                             If mrsPersons!id = UserInfo.id Then
                                lngFindCboIndex = cboEdit.NewIndex
                            End If
                        Else
                            If strText = mrsPersons!���� Then
                                lngFindCboIndex = cboEdit.NewIndex
                            End If
                        End If
                        mrsPersons.MoveNext
                    Loop
                    
                    If cboEdit.ListCount > 0 And cboEdit.ListIndex = -1 Then
                        If lngFindCboIndex <> -1 Then
                            cboEdit.ListIndex = lngFindCboIndex
                        ElseIf strText <> "" Then
                            gobjComlib.cbo.SetText cboEdit, strText
                        Else
                            cboEdit.ListIndex = 0
                        End If
                    End If
                    If mlngModul = pҽ������վ Then
                        If Val(gobjDatabase.GetPara(51, 100)) = 1 Then
                            cboEdit.locked = True
                        End If
                    End If
                End If
                picCbo.Left = CellRect.Left
                picCbo.Top = CellRect.Top
                picCbo.Width = CellRect.Right
                picCbo.Height = CellRect.Bottom
                If cboEdit.locked = False Then
                    cboEdit.Width = picCbo.Width + 30
                Else
                    cboEdit.Width = picCbo.Width + 300
                End If
                gobjControl.CboSetHeight cboEdit, picCbo.Height + 30
                cboEdit.Left = -15
                cboEdit.Top = -15
                '����չ�����
                lngLegth = 0
                For i = 0 To cboEdit.ListCount - 1
                    If lngLegth < LenB(StrConv(cboEdit.List(i), vbFromUnicode)) Then
                        lngLegth = LenB(StrConv(cboEdit.List(i), vbFromUnicode))
                    End If
                Next i
                If lngLegth * picText.TextWidth("1") + 60 > picCbo.Width Then
                    Call gobjControl.CboSetWidth(cboEdit.hWnd, lngLegth * picText.TextWidth("1") + 60)
                Else
                    Call gobjControl.CboSetWidth(cboEdit.hWnd, picCbo.Width)
                End If
                picCbo.BackColor = .BackColor
                picCbo.BorderStyle = 0
                picCbo.Visible = True
                picCbo.ZOrder 0
                picCbo.SetFocus
                picCbo.Tag = .ColKey(.Col)
        End Select
    End With
End Sub

Private Function GetDataToPersons(Optional ByVal strIn As String = "", Optional ByVal blnRetrunSQL As Boolean, Optional strRetrunSQL As String) As ADODB.Recordset
'������Ӧ���ҵ�ҽ����Ա��Ϣ
    Dim strSQL As String, strNewSQL As String, strWhere As String
    Dim blnYn As Boolean
    
    On Error GoTo ErrHand
    If strIn <> "" Then blnYn = True
    
    'ҽ����վ��ֻ��ִ��������Ŀ���ܿ����������ҵ�ҽ�����ٴ���ʿվ���ܻ����ȫԱ���˵�Ȩ�ޣ���Ҫ���ϲ���Ա����(���ڲ����������ˣ�Ҫô�������Ǹò����ģ�Ҫô���Ǿ���ȫԱ����Ȩ��)
    If InStr(mstrPrivs, "ִ��������Ŀ") > 0 Or Not (mlngModul = pҽ������վ) Then
        strNewSQL = " Union " & vbNewLine & _
                            " Select " & UserInfo.id & " id,'" & UserInfo.��� & "' ���,'" & UserInfo.���� & "' ����,'" & UserInfo.���� & "' ���� From Dual "
    End If
        
    If Not mlngModul = pҽ������վ Then
        strWhere = "  Exists (Select 1 From ��Ա����˵�� Where ��Աid = a.Id And Instr(',ҽ��,��ʿ,', ',' || ��Ա���� || ',', 1) <> 0)"
    End If
    
    '��ǰ��¼����Ա������ʾ��ǰ��
    If strNewSQL = "" Then
        strSQL = "Select a.Id, a.���, a.����, a.����" & vbNewLine & _
            " From ��Ա�� a, ������Ա b" & vbNewLine & _
            " Where a.Id = b.��Աid And b.����id = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And" & vbNewLine & _
            "      (a.վ�� = ' & gstrNodeNo & ' Or a.վ�� Is Null) " & vbNewLine & _
            IIf(blnYn, " And (A.��� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & vbNewLine & _
            IIf(strWhere = "", "", " And " & strWhere) & " Order by Decode(a.id," & IIf(blnYn = True, "[4]", "[2]") & ",0,1),a.���"
    Else
        strSQL = "Select a.Id, a.���, a.����, a.����" & vbNewLine & _
            " From ��Ա�� a, ������Ա b" & vbNewLine & _
            " Where a.Id = b.��Աid And b.����id = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And" & vbNewLine & _
            "      (a.վ�� = ' & gstrNodeNo & ' Or a.վ�� Is Null)" & IIf(strWhere = "", "", " And " & strWhere) & vbNewLine & _
            strNewSQL
        If blnYn Then
            strSQL = " Select a.Id, a.���, a.����, a.���� From (" & strSQL & ") a" & vbNewLine & _
                " Where (A.��� Like [2] Or A.���� Like [3] Or A.���� Like [3])  Order by Decode(a.id,[4],0,1),a.���"
        Else
            strSQL = " Select a.Id, a.���, a.����, a.���� From (" & strSQL & ") a Order by Decode(a.id,[2],0,1),a.���"
        End If
    End If
    If blnRetrunSQL Then
        strRetrunSQL = strSQL
    Else
        If blnYn = True Then
            Set GetDataToPersons = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, UCase(strIn) & "%", gstrLike & UCase(strIn) & "%", UserInfo.id)
        Else
            Set GetDataToPersons = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, UserInfo.id)
        End If
    End If
    
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MoveNextCell(Optional ByVal blnNext As Boolean = True, Optional ByVal blnNoMove As Boolean = False, Optional ByVal blnKeyReturn As Boolean) As Boolean
'���ܣ����ݸ�ֵ��У�鴦��
    Dim strText As String
    Dim strMsg As String
    Dim intRow As Integer
    Dim blnFind As Boolean, int��� As Integer
    Dim vPoint As RECT, strName As String
    If mblnAcTive = True Then Exit Function
    On Error GoTo ErrHand
    If mintType >= 0 Then
        Select Case mintType
            Case 0
                If mskʱ��.Text = "____-__-__ __:__" Then
                    strText = ""
                ElseIf Not IsDate(mskʱ��.Text) Then
                    strMsg = "[" & picDate.Tag & "]������Ч��ʱ���ʽ��"
                    Call ShowMsg(strMsg, vbRed)
                    If picDate.Enabled And picDate.Visible Then picDate.SetFocus
'                    If IsDate(cmdDate.Tag) Then
'                        mskʱ��.Text = cmdDate.Tag
'                    Else
'                        mskʱ��.Text = "____-__-__ __:__"
'                    End If
                    Exit Function
                Else
                    strText = mskʱ��.Text
                End If
            Case 1
                strText = lstSelect.Text
            Case 2
                strText = TxtEdit.Text
                If CheckVitalSigns(strText, strMsg) = False Then
                    Call ShowMsg(strMsg, vbRed)
                    If picText.Enabled And picText.Visible Then picText.SetFocus
'                    strText = TxtEdit.Tag
                    Exit Function
                End If
            Case 3
                strText = cboEdit.Text
                If strText <> "" Then
                    Select Case picCbo.Tag
                        Case "����"
                            If InStr(1, ",����,��ѹ,", "," & strText & ",") = 0 And Not IsNumeric(strText) Then
                                strMsg = "¼���[����]���Ƿ��������͵ġ����١���ѹ�����������ͣ�"
                                Call ShowMsg(strMsg, vbRed)
                                If picCbo.Enabled And picCbo.Visible Then picCbo.SetFocus
    '                            strText = cboEdit.Tag
                                Exit Function
                            End If
                        Case "ִ����"
                            blnFind = False
                            mrsPersons.Filter = ""
                            Do While Not mrsPersons.EOF
                                If mrsPersons!���� & "" = strText Then
                                    blnFind = True
                                    Exit Do
                                End If
                                mrsPersons.MoveNext
                            Loop
                            If blnFind = False Then
                                If blnKeyReturn = True Then
                                    vPoint.Left = picCbo.Left + vsfExec.Left
                                    vPoint.Top = picCbo.Top + vsfExec.Top + picCbo.Height
                                    blnFind = FindPerson(strText, vPoint.Left, vPoint.Top, strName)
                                    If blnFind = True And strName <> "" Then
                                        strText = strName
                                    End If
                                End If
                            End If
                            If blnFind = False Then
                                strMsg = "¼���[ִ����]������Ч��ִ���˷�Χ�ڣ�"
                                Call ShowMsg(strMsg, vbRed)
                                If picCbo.Enabled And picCbo.Visible Then picCbo.SetFocus
    '                            strText = cboEdit.Tag
                                Exit Function
                            End If
                    End Select
                End If
        End Select
        If vsfExec.TextMatrix(vsfExec.Row, vsfExec.Col) <> strText Then
            If mblnChange = False Then mblnChange = True
            Call ChangeDataState
        End If
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.Col) = strText
        Call HiddenEditControl
    End If
    If (vsfExec.Col = vsfExec.ColIndex("ִ��ʱ��") Or vsfExec.Col = vsfExec.ColIndex("ִ����")) And Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("��¼����"))) = 2 Then
        If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ִ��ʱ��")) <> "" And vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ִ����")) <> "" Then
            int��� = Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("���")))
            If vsfExec.Row + 1 < vsfExec.Rows Then
                If Val(vsfExec.TextMatrix(vsfExec.Row + 1, vsfExec.ColIndex("��¼����"))) <> 2 Then
                    vsfExec.Rows = vsfExec.Rows + 1
                    vsfExec.TextMatrix(vsfExec.Rows - 1, vsfExec.ColIndex("��¼����")) = 2
                    vsfExec.TextMatrix(vsfExec.Rows - 1, vsfExec.ColIndex("���")) = int��� + 1
                    vsfExec.TextMatrix(vsfExec.Rows - 1, 0) = "��ע����"
                    vsfExec.TextMatrix(vsfExec.Rows - 1, 1) = int��� + 1 & "Сʱ"
                    vsfExec.MergeRow(vsfExec.Rows - 1) = True
                    vsfExec.RowPosition(vsfExec.Rows - 1) = vsfExec.Rows - 3
                    vsfExec.Cell(flexcpAlignment, vsfExec.FixedRows, vsfExec.FixedCols, vsfExec.Rows - 1, vsfExec.Cols - 1) = flexAlignCenterCenter
                End If
            End If
        End If
    End If
    MoveNextCell = True
    If blnNoMove = True Then Exit Function
    If blnNext Then
toMoveNextCol:
        '������һ��
        If vsfExec.Col < mlngNoEditor - 1 Then
            vsfExec.Col = vsfExec.Col + 1
            If vsfExec.ColWidth(vsfExec.Col) = 0 Or vsfExec.ColHidden(vsfExec.Col) Or mintType = -1 Then GoTo toMoveNextCol
        Else
toMoveNextRow:
            '������һ��
            mblnShow = False
            If vsfExec.Row + 1 < vsfExec.Rows Then
                vsfExec.Row = vsfExec.Row + 1
            End If
            If vsfExec.RowHidden(vsfExec.Row) Then
                If vsfExec.Row < vsfExec.Rows - 1 Then
                    If txtִ��ժҪ.Enabled And txtִ��ժҪ.Visible Then txtִ��ժҪ.SetFocus
                Else
                    For intRow = vsfExec.Rows - 1 To vsfExec.FixedRows Step -1
                        If vsfExec.RowHidden(intRow) = False Then
                            vsfExec.Row = intRow
                            Exit For
                        End If
                    Next intRow
                End If
            End If
            mblnShow = True
            vsfExec.Col = vsfExec.ColIndex("ִ��ʱ��")
        End If
    Else
toMovePrevCol:
        If vsfExec.Col > vsfExec.ColIndex("ִ��ʱ��") Then       '�����¼���϶��л�ʿǩ����
            vsfExec.Col = vsfExec.Col - 1
            If vsfExec.ColWidth(vsfExec.Col) = 0 Or vsfExec.ColHidden(vsfExec.Col) Or mintType = -1 Then GoTo toMovePrevCol
        Else
toMovePrevRow:
            '������һ��
            If vsfExec.Row > vsfExec.FixedRows Then
                vsfExec.Row = vsfExec.Row - 1
                If vsfExec.RowHidden(vsfExec.Row) Then GoTo toMovePrevRow
                vsfExec.Col = mlngNoEditor - 1
            End If
        End If
    End If
    If vsfExec.ColIsVisible(vsfExec.Col) = False Then
        vsfExec.LeftCol = vsfExec.Col
    End If
    If vsfExec.RowIsVisible(vsfExec.Row) = False Then
        vsfExec.TopRow = vsfExec.Row
    End If
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ChangeDataState()
'���ܣ������ݷ����仯���ô˹��̣��޸�����״̬
    Select Case Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("״̬")))
        '1,2,3��ʾ��ԭ�е����ݽ��в�����1��ԭʼ��2-�޸�;3-ɾ��
        '0,4 ��ʾ�Խ�Ҫ���������ݲ�����0-��������4-����
        Case 3
            If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ִ��ʱ��")) <> "" Then vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("״̬")) = 2
        Case 1, 2
            If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ִ��ʱ��")) <> "" Then
                vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("״̬")) = 2
            Else
                vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("״̬")) = 3
            End If
        Case 0 '������
            If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ִ��ʱ��")) <> "" Then vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("״̬")) = 4
        Case 4
            If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ִ��ʱ��")) = "" Then vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("״̬")) = 0
    End Select
End Sub

Private Function FindPerson(ByVal strText As String, ByVal lngLeft As Long, ByVal lngTop As Long, strName As String) As Boolean
    Dim rsUser As ADODB.Recordset
    Dim strSQL As String
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHand
    If strText <> "" Then
        Call GetDataToPersons(strText, True, strSQL)
        Set rsUser = gobjDatabase.ShowSQLSelect(Me, strSQL, 0, "", False, strText, "��ѡ����Ա", False, False, True, lngLeft, lngTop, 0, blnCancel, False, False, _
                    mlng����ID, UCase(strText) & "%", gstrLike & UCase(strText) & "%", UserInfo.id)
        If Not rsUser Is Nothing Then
            If blnCancel = False Then
                If rsUser.EOF Then Exit Function
                strName = Nvl(rsUser!����)
            End If
        Else
            Exit Function
        End If
    End If
    FindPerson = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckVitalSigns(strText As String, strMsg As String) As Boolean
'���ܣ������������ݺϷ��Լ��
    Dim arrData, arrName
    Dim i As Integer
    Dim strֵ�� As String, intС�� As Integer, int���� As Integer
    Dim dblMin As Double, dblMax As Double
    Dim blnMatch As Boolean
    
    On Error GoTo ErrHand
    arrData = Array()
    arrName = Array()
    If strText <> "" Then
        If picText.Tag = "Ѫѹ" Then
            If InStr(1, strText, "/") = 0 Then
                strMsg = "¼���[" & picText.Tag & "]��ʽ����ȷ����ȷ��ʽ������ѹ/����ѹ��"
                Exit Function
            End If
            ReDim Preserve arrData(UBound(arrData) + 1)
            arrData(UBound(arrData)) = Mid(strText, 1, InStr(1, strText, "/") - 1)
            ReDim Preserve arrName(UBound(arrName) + 1)
            arrName(UBound(arrName)) = "����ѹ"
            
            ReDim Preserve arrData(UBound(arrData) + 1)
            arrData(UBound(arrData)) = Mid(strText, InStr(1, strText, "/") + 1)
            ReDim Preserve arrName(UBound(arrName) + 1)
            arrName(UBound(arrName)) = "����ѹ"
        Else
            ReDim Preserve arrData(UBound(arrData) + 1)
            arrData(UBound(arrData)) = strText
            ReDim Preserve arrName(UBound(arrName) + 1)
            arrName(UBound(arrName)) = picText.Tag
        End If
        For i = 0 To UBound(arrName)
            If Not IsNumeric(arrData(i)) Then
                strMsg = "¼���[" & arrName(i) & "]������Ч�����ָ�ʽ��"
                Exit Function
            End If
            'ֵ��Χ���
            strֵ�� = "": intС�� = -1: int���� = -1
            mrsItems.Filter = "������='" & arrName(i) & "'"
            If Not mrsItems.EOF Then
                Select Case CStr(arrName(i))
                    Case "����"
                        blnMatch = mrsItems!��λ & "" = "��"
                    Case "����", "����"
                        blnMatch = mrsItems!��λ & "" = "��/��"
                    Case "����ѹ", "����ѹ"
                        blnMatch = mrsItems!��λ & "" = "mmHg"
                End Select
                If blnMatch = True Then
                    strֵ�� = mrsItems!��ֵ�� & ""
                    intС�� = Val(mrsItems!С�� & "")
                    int���� = Val(mrsItems!���� & "")
                End If
            End If
            If InStr(1, strֵ��, ";") = 0 Or intС�� = -1 Or int���� - 1 Then
             '�Ҳ�����ʹ��ȱʡֵ
                Select Case CStr(arrName(i))
                    Case "����"
                        If InStr(1, strֵ��, ";") = 0 Then strֵ�� = "35;42"
                        If intС�� = -1 Then intС�� = 1
                        If int���� = -1 Then int���� = 4
                    Case "����"
                        If InStr(1, strֵ��, ";") = 0 Then strֵ�� = "20;300"
                        If intС�� = -1 Then intС�� = 0
                        If int���� = -1 Then int���� = 3
                    Case "����"
                        If InStr(1, strֵ��, ";") = 0 Then strֵ�� = "15;50"
                        If intС�� = -1 Then intС�� = 0
                        If int���� = -1 Then int���� = 2
                    Case "����ѹ", "����ѹ"
                        If InStr(1, strֵ��, ";") = 0 Then strֵ�� = "50;190"
                        If intС�� = -1 Then intС�� = 0:
                        If int���� = -1 Then int���� = 3
                End Select
            End If
            strText = arrData(i)
            '���ȼ��
            If Len(strText) > int���� Then
                strMsg = "¼���[" & arrName(i) & "]���ݳ�������󳤶ȣ�" & int���� & "��"
                Exit Function
            End If
                    
            If intС�� <> 0 Then
                If InStr(1, strText, ".") <> 0 Then
                    strText = Mid(strText, InStr(1, strText, ".") + 1)
                    If Len(strText) > intС�� Then
                        strMsg = "¼���[" & arrName(i) & "]¼��С�����ֳ����˺Ϸ�����" & intС�� & "λ��"
                        Exit Function
                    End If
                End If
            End If
            strText = arrData(i)
            If strֵ�� <> "" Then
                dblMin = Val(Split(strֵ��, ";")(0))
                dblMax = Val(Split(strֵ��, ";")(1))
                If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                    strMsg = "¼���[" & arrName(i) & "]���ݲ���" & Format(dblMin, "#0.00") & "��" & Format(dblMax, "#0.00") & "����Ч��Χ��"
                    Exit Function
                End If
            End If
            If Val(strText) < 1 And Val(strText) > 0 Then strText = "0" & Val(strText)
            arrData(i) = strText
        Next
        If picText.Tag = "Ѫѹ" Then
            strText = arrData(0) & "/" & arrData(1)
        Else
            strText = arrData(0)
        End If
    End If
    CheckVitalSigns = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitExecBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
   Dim cbrCustom As CommandBarControlCustom
   
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsExec.VisualTheme = xtpThemeOfficeXP
    With Me.cbsExec.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
    End With
    Set cbsExec.Icons = gobjCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
    
    Set objBar = cbsExec.Add("������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "�˶�"): objControl.ToolTipText = "��Ѫǰ�˶�"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit, "ȡ��"): objControl.ToolTipText = "ȡ����Ѫǰ�˶�"
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Sign, "ǩ��"): objControl.ToolTipText = "ǩ������": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��"): objControl.ToolTipText = "ȡ��ǩ������"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Clear, "���"): objControl.ToolTipText = "���������": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.BeginGroup = True
        If mblnOnlyRead = False Then
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.Flags = xtpFlagRightAlign
            cbrCustom.Handle = picinfo.hWnd
        End If
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsExec.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add FCONTROL, Asc("C"), conMenu_Edit_Transf_Cancle
        .Add FCONTROL, Asc("E"), conMenu_File_Exit
    End With
End Sub

Private Sub ShowMsg(ByVal strText As String, Optional lngColor As Long = vbBlack)
    lblPrompt.Caption = strText
    lblPrompt.ForeColor = lngColor
End Sub

Private Function SetMessages(ByRef arrSQL As Variant, Optional ByVal blnRead As Boolean = False) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lng����ID As Long, lng����id As Long, lng����ID As Long
    Dim lng����id As Long
    Dim int������Դ As Integer
    Dim str���Ѳ��� As String
    Dim bln��Ѫ��Ӧ As Boolean, intRow As Integer
    
    On Error GoTo ErrHand
    arrSQL = Array()
    strSQL = "select ��ҳid,�Һŵ�,����id,���˿���id,������Դ from ����ҽ����¼ where id = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng���ID)
    
    If rsTmp.State = adStateClosed Then Exit Function
    If rsTmp.RecordCount = 0 Then Exit Function
    lng����ID = Val(rsTmp!����id)
    lng����id = Val(rsTmp!���˿���id)
    int������Դ = Val(rsTmp!������Դ)
    If int������Դ = 2 Then
        lng����id = Val(rsTmp!��ҳid)
        strSQL = "select ��ǰ����id from ������ҳ where ����id = [1] and ��ҳid = [2]  "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng����id)
        lng����ID = Val(rsTmp!��ǰ����id)
    Else
        lng����ID = Val(rsTmp!���˿���id)
        strSQL = "select id �Һ�id from ���˹Һż�¼ where no = [1] and ����id = [2] "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTmp!�Һŵ� & "", Val(rsTmp!����id))
        lng����id = Val(rsTmp!�Һ�ID)
    End If
        
    If blnRead = False Then
        strSQL = "select ID,���ͱ���,ҵ���ʶ from ҵ����Ϣ�嵥 where ����ID = [1] and ����id = [2] and �Ƿ����� = 0 "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng����id)
        
        'ȷ��Ҫ���ѵĲ���
        str���Ѳ��� = IIf(Val(lng����id) = 0, "", lng����id)
        If lng����ID <> lng����id Then
            If str���Ѳ��� = "" Then
                str���Ѳ��� = IIf(lng����ID = 0, "", lng����ID)
            Else
                str���Ѳ��� = str���Ѳ��� & IIf(lng����ID = 0, "", "," & lng����ID)
            End If
        End If
        '��ѯ�Ƿ���ڱ�ҽ����Ѫ����Ѫ��Ӧ��Ϣ
        rsTmp.Filter = "���ͱ��� = 'ZLHIS_BLOOD_006' And ҵ���ʶ = '" & mlng���ID & ":" & mlng�շ�ID & "'"
        '�Ƿ�����Ѫ��Ӧ
        With vsfExec
            For intRow = .FixedRows To .Rows - 1
                If IsDate(.TextMatrix(intRow, .ColIndex("ִ��ʱ��"))) Then
                    If .TextMatrix(intRow, .ColIndex("��Ѫ��Ӧ")) <> "" And .TextMatrix(intRow, .ColIndex("��Ѫ��Ӧ")) <> "��" Then
                        bln��Ѫ��Ӧ = True
                        Exit For
                    End If
                End If
            Next
        End With
        If bln��Ѫ��Ӧ = True Then
            If rsTmp.RecordCount = 0 Then
                strSQL = "Zl_ҵ����Ϣ�嵥_Insert(" & lng����ID & "," & lng����id & ","  '����id ����id
                strSQL = strSQL & Val(lng����id) & ","     '�������id
                strSQL = strSQL & Val(lng����ID) & ","      '���ﲡ��id
                strSQL = strSQL & int������Դ & ","                                      '������Դ
                strSQL = strSQL & "'������Ѫ��Ӧ���뼰ʱ��д��Ѫ��Ӧ����','"             '��Ϣ����
                strSQL = strSQL & IIf(Val(int������Դ) = 1, "1000", "0100") & "','ZLHIS_BLOOD_006',"     ' ���ѳ���, ���ͱ���
                strSQL = strSQL & "'" & mlng���ID & ":" & mlng�շ�ID & "',"                      'ҵ���ʶ�����id:�շ�id��
                strSQL = strSQL & "1,0,NULL,'" & str���Ѳ��� & "',NULL)"                                                   '���ȳ̶ȣ��Ƿ����ģ��Ǽ�ʱ��,���Ѳ���
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        Else    '����Ѫ��Ӧ�����ѯ�Ƿ�����з�Ӧ��Ϣ�����У���Ϊ�Ѷ���
            If rsTmp.RecordCount > 0 Then
                strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����id & ",'ZLHIS_BLOOD_006',"
                strSQL = strSQL & "3,'" & UserInfo.���� & "'," & lng����ID & ",NULL,"
                strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        End If
    
        rsTmp.Filter = "���ͱ��� = 'ZLHIS_BLOOD_007' And ҵ���ʶ = '" & mlng���ID & ":" & mlngҽ��ID & ":" & mlng�շ�ID & "'"
        If IsDate(vsfExec.TextMatrix(vsfExec.Rows - 1, vsfExec.ColIndex("ִ��ʱ��"))) Then '����ִ��
            If rsTmp.RecordCount = 0 Then
                strSQL = "Zl_ҵ����Ϣ�嵥_Insert(" & lng����ID & "," & lng����id & ","  '����id ����id
                strSQL = strSQL & Val(lng����id) & ","      '�������id
                strSQL = strSQL & Val(lng����ID) & ","      '���ﲡ��id
                strSQL = strSQL & int������Դ & ","                                      '������Դ
                strSQL = strSQL & "'��Ѫ��ɣ�����24Сʱ���ջ�Ѫ����','"                         '��Ϣ����
                strSQL = strSQL & IIf(Val(int������Դ) = 1, "0001", "0010") & "','ZLHIS_BLOOD_007',"     ' ���ѳ���, ���ͱ���
                strSQL = strSQL & "'" & mlng���ID & ":" & mlngҽ��ID & ":" & mlng�շ�ID & "',"                      'ҵ���ʶ�����id:�շ�id��
                strSQL = strSQL & "1,0,NULL,'" & str���Ѳ��� & "',NULL)"                                                   '���ȳ̶ȣ��Ƿ����ģ��Ǽ�ʱ��,���Ѳ���                                                      '
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        Else    '
            If rsTmp.RecordCount > 0 Then
                strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����id & ",'ZLHIS_BLOOD_007',"
                strSQL = strSQL & IIf(Val(int������Դ) = 1, 4, 3) & ",'" & UserInfo.���� & "'," & lng����ID & ",NULL,"
                strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        End If
    Else
        '�Ƿ���ڸ�Ѫ����Ϣ����������Ϊ�Ѷ�
        strSQL = "Select a.Id, a.���ͱ���, a.����id, a.ҵ���ʶ" & vbNewLine & _
                    "From ҵ����Ϣ�嵥 a" & vbNewLine & _
                    "Where a.����id = [1] And a.����id = [2] And a.�Ƿ����� = 0 And a.���ͱ��� In ('ZLHIS_BLOOD_006', 'ZLHIS_BLOOD_007')"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "Ѫ�������Ϣ", lng����ID, lng����id)
        
        For i = 0 To 1
            '��ZLHIS_BLOOD_006����Ϣ��Ϊ�Ѷ�
            If i = 0 Then rsTmp.Filter = "ҵ���ʶ = '" & mlng���ID & ":" & mlng�շ�ID & "'"
            '��ZLHIS_BLOOD_007����Ϣ��Ϊ�Ѷ�
            If i = 1 Then rsTmp.Filter = "ҵ���ʶ = '" & mlng���ID & ":" & mlngҽ��ID & ":" & mlng�շ�ID & "'"
            If Not rsTmp.EOF Then
                rsTmp.MoveFirst
                Do While Not rsTmp.EOF
                    strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & rsTmp!����id & ",'" & rsTmp!���ͱ��� & "',"
                    strSQL = strSQL & IIf(Val(int������Դ) = 1, 4, 3) & ",'" & UserInfo.���� & "'," & lng����ID & ",NULL,"
                    strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    rsTmp.MoveNext
                Loop
            End If
        Next
    End If
    SetMessages = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SignData(Optional ByVal blnǩ�� As Boolean = False)
    Dim strName As String, strǩ���� As String
    Dim blnSign  As Boolean, strSQL As String
    Dim int��¼���� As Integer, int��� As Integer
    
    On Error GoTo ErrHand
    int��¼���� = Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("��¼����")))
    int��� = Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("���")))
    If blnǩ�� = True Then '����
        If Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("״̬"))) <> 1 Then
            MsgBox "�뱣�����ݺ���ǩ���ˣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If IsDate(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ִ��ʱ��"))) = False Or Trim(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ִ����"))) = "" Then
            MsgBox "ִ��ʱ���ִ���˲���Ϊ�գ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strSQL = "Zl_ѪҺִ�м�¼_Sign(" & mlng�շ�ID & "," & int��¼���� & "," & int��� & ",'" & UserInfo.���� & "',1)"
        Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ����")) = UserInfo.����
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ��ʱ��")) = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    Else
        'ȡ���Ǽ���Ƿ��ǵ�ǰ����Ա����������Ҫ���������֤
        strName = vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ����"))
        If strName <> UserInfo.���� Then
            strǩ���� = gobjDatabase.UserIdentifyByUser(Me, "�Ǳ���ȡ������������ǩ���˵��û�����������������֤��", 100, mlngModul, "ִ������Ǽ�", , True)
            If strǩ���� = "" Then Exit Sub
            If strǩ���� <> strName Then
                MsgBox "ֻ��ȡ���Լ�ǩ���ļ�¼����ǰǩ������""" & strName & """", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '����ǩ��
        strSQL = "Zl_ѪҺִ�м�¼_Sign(" & mlng�շ�ID & "," & int��¼���� & "," & int��� & ",NULL,0)"
        Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ����")) = ""
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ��ʱ��")) = ""
    End If
'    vsfExec.Cell(flexcpForeColor, vsfExec.Row, vsfExec.FixedCols, vsfExec.Row, vsfExec.Cols - 1) = IIf(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("ǩ����")) <> "", vbRed, vbBlack)
    Call vsfExec_AfterRowColChange(0, 0, vsfExec.Row, vsfExec.Col)
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function AutoAdviceFinish() As Boolean
'���ܣ��жϸ�ҽ�������ѪҺ�Ƿ��Ѿ����ִ�У������������Զ����ҽ��ִ��
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    strSQL = "Select a.�շ�ID from ѪҺ���ͼ�¼ A,ѪҺ���ͼ�¼ B where a.�䷢ID=B.�䷢ID and B.�շ�ID=[1] and nvl(a.ִ��״̬,0) not in (2,3)"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "�ж�����ѪҺ�Ƿ��Ѿ����ִ��", mlng�շ�ID)
    If rsTmp.RecordCount = 0 Then
        '�Զ���ҽ�����Ϊ���
        strSQL = "ZL_����ҽ��ִ��_Finish(" & mlngҽ��ID & "," & mlng���ͺ� & ",Null,0,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
        AutoAdviceFinish = True
    End If
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function
