VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmOpenStopedPlanBySN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ͣ���Դ"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpenStopedPlanBySN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleMode       =   0  'User
   ScaleWidth      =   8891.551
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSignalSourceSelect 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   4935
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1170
      Width           =   4935
      Begin VB.ComboBox cboDept 
         Height          =   330
         Left            =   525
         TabIndex        =   6
         Top             =   50
         Width           =   1785
      End
      Begin VB.ComboBox cboDoctor 
         Height          =   330
         ItemData        =   "frmOpenStopedPlanBySN.frx":000C
         Left            =   2985
         List            =   "frmOpenStopedPlanBySN.frx":000E
         TabIndex        =   7
         Top             =   50
         Width           =   1785
      End
      Begin VB.Label lblDeptFilter 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   60
         TabIndex        =   30
         Top             =   110
         Width           =   420
      End
      Begin VB.Label lblDoctorFilter 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         Height          =   210
         Left            =   2520
         TabIndex        =   29
         Top             =   105
         Width           =   420
      End
   End
   Begin VB.Frame fraSplitY 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   4440
      TabIndex        =   38
      Top             =   6570
      Width           =   2385
   End
   Begin VB.Frame fraRecordInfo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3570
      TabIndex        =   34
      Top             =   2190
      Width           =   7005
      Begin VB.TextBox txtͣ��ʱ�� 
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   0
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   50
         Width           =   1815
      End
      Begin VB.TextBox txt�޺��� 
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   50
         Width           =   1005
      End
      Begin VB.TextBox txt��Լ�� 
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   0
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   50
         Width           =   1005
      End
      Begin VB.Label lblͣ��ʱ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ͣ��ʱ��"
         Height          =   210
         Index           =   0
         Left            =   4290
         TabIndex        =   37
         Top             =   110
         Width           =   840
      End
      Begin VB.Label lbl��Լ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Լ��"
         Height          =   210
         Index           =   0
         Left            =   2280
         TabIndex        =   36
         Top             =   110
         Width           =   630
      End
      Begin VB.Label lbl�޺��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޺���"
         Height          =   210
         Index           =   0
         Left            =   300
         TabIndex        =   35
         Top             =   110
         Width           =   630
      End
   End
   Begin VB.Frame fraSplit 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3750
      MousePointer    =   9  'Size W E
      TabIndex        =   33
      Top             =   6420
      Width           =   67
   End
   Begin VB.PictureBox picVisitDate 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6210
      ScaleHeight     =   405
      ScaleWidth      =   4725
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1350
      Width           =   4725
      Begin VB.TextBox txtWorkTime 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   50
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Left            =   930
         TabIndex        =   9
         Top             =   50
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   160104451
         CurrentDate     =   42713
      End
      Begin VB.Label lblWorkTime 
         AutoSize        =   -1  'True
         Caption         =   "�ϰ�ʱ��"
         Height          =   210
         Left            =   2610
         TabIndex        =   32
         Top             =   110
         Width           =   840
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   210
         Left            =   30
         TabIndex        =   31
         Top             =   110
         Width           =   840
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSignalSource 
      Height          =   4845
      Left            =   0
      TabIndex        =   11
      Top             =   2070
      Width           =   3465
      _cx             =   6112
      _cy             =   8546
      Appearance      =   2
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOpenStopedPlanBySN.frx":0010
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
   Begin VB.Frame fraSignalSource 
      Caption         =   "��Դ��Ϣ"
      Height          =   1095
      Left            =   0
      TabIndex        =   22
      Top             =   30
      Width           =   11895
      Begin VB.TextBox txt���� 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   285
         Width           =   1275
      End
      Begin VB.TextBox txtSignalNO 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   285
         Width           =   1035
      End
      Begin VB.TextBox txtDoctor 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   8910
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   285
         Width           =   2175
      End
      Begin VB.TextBox txtDept 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   285
         Width           =   2625
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   675
         Width           =   3045
      End
      Begin VB.Label lblSignalNO 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   300
         TabIndex        =   27
         Top             =   345
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   2070
         TabIndex        =   26
         Top             =   345
         Width           =   420
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ"
         Height          =   210
         Left            =   300
         TabIndex        =   25
         Top             =   735
         Width           =   420
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   4710
         TabIndex        =   24
         Top             =   345
         Width           =   420
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         Height          =   210
         Left            =   8460
         TabIndex        =   23
         Top             =   345
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   380
      Left            =   9555
      TabIndex        =   21
      Top             =   6990
      Width           =   1250
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   405
      Left            =   3570
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1710
      Width           =   2385
      _Version        =   589884
      _ExtentX        =   4207
      _ExtentY        =   714
      _StockProps     =   64
   End
   Begin VB.PictureBox picTimeWork 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   0
      Left            =   3570
      ScaleHeight     =   4065
      ScaleWidth      =   8385
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Width           =   8385
      Begin VB.TextBox txtOpen 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   930
         TabIndex        =   18
         Text            =   "0"
         Top             =   3690
         Width           =   1170
      End
      Begin MSComCtl2.UpDown updOpen 
         Height          =   330
         Index           =   0
         Left            =   2130
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3690
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         BuddyControl    =   "txtOpen(0)"
         BuddyDispid     =   196641
         BuddyIndex      =   0
         OrigLeft        =   2355
         OrigTop         =   3690
         OrigRight       =   2610
         OrigBottom      =   4020
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfTimeWork 
         Height          =   3210
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   420
         Width           =   8220
         _cx             =   14499
         _cy             =   5662
         Appearance      =   2
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
         GridColor       =   12632256
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOpenStopedPlanBySN.frx":00EC
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
         PicturesOver    =   -1  'True
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
      Begin VB.Label lblToolTip 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   0
         Left            =   2460
         TabIndex        =   39
         Top             =   3750
         Width           =   420
      End
      Begin VB.Label lblOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   28
         Top             =   3750
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   380
      Left            =   8250
      TabIndex        =   20
      Top             =   6990
      Width           =   1250
   End
End
Attribute VB_Name = "frmOpenStopedPlanBySN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ���
Private mlngModule As Long
Private mlngDeptID As Long, mlngDoctorID As Long
Private mlng��¼ID As Long '��ԴID

Private mblnOK As Boolean, mblnFirst As Boolean
Private msngStartX As Single    '�ƶ�ǰ����λ��

'�Һ����״̬��0-���һ��ԤԼ�ĺ�;1-�ѹ�;2-�Ѿ�ԤԼ;3-Ԥ����;4-�Ѿ��˺�;5-�Ѿ�����;6-��ͣ��
Private Enum SNState
    ���� = 0
    �ѹ� = 1
    ��Լ = 2
    Ԥ�� = 3
    �˺� = 4
    ���� = 5
    ͣ�� = 6
End Enum

Private mrsDept As ADODB.Recordset, mrsDoctor As ADODB.Recordset
Private mrsRecord As ADODB.Recordset, mrsRecordCount As ADODB.Recordset
Private mblnNotChange As Boolean, mblnChanged As Boolean
Private mblnCboClick As Boolean     '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
Private mlngPreRow As Long

Public Function ShowMe(ByVal frmMain As Object, ByVal lngModule As Long, _
    Optional ByVal lng��¼ID As Long, _
    Optional ByVal lngDeptID As Long, _
    Optional ByVal lngDoctorID As Long) As Boolean
    '������ڣ�����ͣ����ţ����ű�������������ſ����ҷ�ʱ�ε�
    '��Σ�
    '   frmMain ���õ�������
    '   lngModule ����ģ���
    '   lng��¼ID ��¼ID,1114ģ�����ʱ����
    '   lngDeptID ����ID
    '   lngDoctorID ҽ��ID
    '˵����
    '   lngModule������1114ʱ��
    '   1.���������ҽ��ID,�����ֻ��ѡ�����Ա��������,ҽ�����ܱ༭
    '   2.��������˿���ID,����Ҳ��ܱ༭
    '   ����ȱʡΪ��ǰ����
    mlngModule = lngModule
    mlngDeptID = lngDeptID: mlngDoctorID = lngDoctorID
    mlng��¼ID = lng��¼ID

    On Error Resume Next
    mblnOK = False
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    ShowMe = mblnOK
End Function

Private Sub cboDept_Click()
    Dim lngDept As Long, lngDoctor As Long
    
    On Error GoTo ErrHandler
    mblnCboClick = True
    If cboDept.ListIndex <> -1 Then
        lngDept = cboDept.ItemData(cboDept.ListIndex)
    End If
    If mlngDoctorID = 0 Then Call FillDoctor(lngDept, mlngDoctorID)
    If cboDoctor.ListIndex <> -1 Then
        lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), lngDept, lngDoctor
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub FillDoctor(Optional lng����id As Long, Optional ByVal lngDefault As Long)
'���ܣ�����ָ���Ŀ�������ID��ȡ����дҽ���б�,ȱʡҽ��
    Dim strOldID As String
    
    On Error GoTo ErrHandler
    cboDoctor.Clear
    If mrsDoctor Is Nothing Then Set mrsDoctor = GetAllDoctor
    mrsDoctor.Filter = ""
    If lng����id <> 0 Then
        mrsDoctor.Filter = "����ID=" & lng����id
    End If
    
    Do While Not mrsDoctor.EOF
        If InStr("," & strOldID & ",", "," & Val(Nvl(mrsDoctor!ID)) & ",") = 0 Then
            cboDoctor.AddItem Nvl(mrsDoctor!����) & "-" & Nvl(mrsDoctor!����)
            cboDoctor.ItemData(cboDoctor.NewIndex) = Val(Nvl(mrsDoctor!ID))
            If lngDefault = Val(Nvl(mrsDoctor!ID)) Then
                cboDoctor.ListIndex = cboDoctor.NewIndex
            End If
            strOldID = strOldID & "," & Val(Nvl(mrsDoctor!ID))
        End If
        mrsDoctor.MoveNext
    Loop
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDept_GotFocus()
    gobjControl.TxtSelAll cboDept
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim lngDoctor As Long
    
    On Error GoTo ErrHandler
    mblnCboClick = False
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mblnCboClick = True
    If Trim(cboDept.Text) = "" Then
        Call FillDoctor(, mlngDoctorID)
        If cboDoctor.ListIndex <> -1 Then
            lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
        End If
        LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), , lngDoctor
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cboDept.ListIndex < 0 Then
        If mrsDept Is Nothing Then Call FillDept(mlngDeptID)
        If zlSelectDept(Me, mlngModule, cboDept, mrsDept, cboDept.Text) = False Then
            KeyAscii = 0: mblnCboClick = False
            Exit Sub
        End If
    Else
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub FillDept(Optional ByVal lngDefault As Long)
    '���ܣ���ȡ�����ؿ����б�,ȱʡ����
    '������lngDefault - ȱʡ����ID
    Dim strSQL As String, strOldID As String
    
    On Error GoTo ErrHandler
    cboDept.Clear
    If mrsDept Is Nothing Then
        Set mrsDept = GetDepartments("�ٴ�", "1,3", mlngDoctorID)
    End If
    mrsDept.Filter = ""
    Do While Not mrsDept.EOF
        If InStr("," & strOldID & ",", "," & Val(Nvl(mrsDept!ID)) & ",") = 0 Then 'һ�����ſ���ͬʱ���ڲ��ƺ��ٴ�,��������ͬ��
            cboDept.AddItem Nvl(mrsDept!����)
            cboDept.ItemData(cboDept.NewIndex) = Val(Nvl(mrsDept!ID))
            If lngDefault = Val(Nvl(mrsDept!ID)) Then
                cboDept.ListIndex = cboDept.NewIndex
            End If
            strOldID = strOldID & "," & Val(Nvl(mrsDept!ID))
        End If
        mrsDept.MoveNext
    Loop
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
    Dim Index As Integer
    
    On Error GoTo ErrHandler
    If cboDept.Text = "" Then
        cboDept.ListIndex = -1
    Else
        Index = SeekCboIndex(cboDept, NeedName(cboDept.Text))
        If Index = -1 Then
            cboDept.ListIndex = -1: cboDept.Text = ""
        ElseIf cboDept.ListIndex <> Index Then
            cboDept.ListIndex = Index
        End If
    End If
    If mblnCboClick = False Then cboDept_Click
    mblnCboClick = False
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDoctor_Click()
    Dim lngDept As Long, lngDoctor As Long
    
    On Error GoTo ErrHandler
    mblnCboClick = True
    If cboDept.ListIndex <> -1 Then
        lngDept = cboDept.ItemData(cboDept.ListIndex)
    End If
    If cboDoctor.ListIndex <> -1 Then
        lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), lngDept, lngDoctor
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDoctor_GotFocus()
    gobjControl.TxtSelAll cboDoctor
End Sub

Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngDept As Long
    
    On Error GoTo ErrHandler
    mblnCboClick = False
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mblnCboClick = True
    If Trim(cboDoctor.Text) = "" Then
        cboDoctor.ListIndex = -1
        If cboDept.ListIndex <> -1 Then
            lngDept = cboDept.ItemData(cboDept.ListIndex)
        End If
        LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), lngDept
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cboDoctor.ListIndex < 0 Then
        If mrsDoctor Is Nothing Then Call FillDoctor(, mlngDoctorID)
        If zlPersonSelect(Me, mlngModule, cboDoctor, mrsDoctor, cboDoctor.Text) = False Then
            KeyAscii = 0: mblnCboClick = False
            Exit Sub
        End If
    Else
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDoctor_Validate(Cancel As Boolean)
    Dim Index As Integer
    
    On Error GoTo ErrHandler
    If cboDoctor.Text = "" Then
        cboDoctor.ListIndex = -1
    Else
        Index = SeekCboIndex(cboDoctor, NeedName(cboDoctor.Text))
        If Index = -1 Then
            cboDoctor.ListIndex = -1: cboDoctor.Text = ""
        ElseIf cboDoctor.ListIndex <> Index Then
            cboDoctor.ListIndex = Index
        End If
    End If
    If mblnCboClick = False Then cboDept_Click
    mblnCboClick = False
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer, cllSql As Collection
    Dim blnTrans As Boolean
    Dim strIDs As String
    
    On Error GoTo ErrHandler
    If mblnChanged = False Then
        MsgBox "����δ�����κΰ��ŵĿ������������ܱ��棡", vbInformation, gstrSysName
        Exit Sub
    End If
    If mlngModule = 1114 Then
        If txtOpen(0).Enabled Then
            strSQL = "Select Nvl(a.�Ƿ���ſ���, 0) * Nvl(a.�Ƿ��ʱ��, 0) As �������ʱ��," & vbNewLine & _
                    "        Decode(a.ͣ�￪ʼʱ��, Null, 0, 1) As ��ͣ��," & vbNewLine & _
                    "        Decode(Sign(Nvl(a.ͣ�￪ʼʱ��, Sysdate) - Sysdate), -1, 1, 0) As ��ʱ" & vbNewLine & _
                    " From �ٴ������¼ A" & vbNewLine & _
                    " Where a.Id = [1]"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��鰲��", mlng��¼ID)
            If rsTemp.EOF Then
                MsgBox "δ�ҵ��������ݣ����ܿ���ͣ�ﰲ�ţ�", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!�������ʱ��)) = 0 Then
                MsgBox "���ﰲ�Ų������������������ʱ�εģ����ܿ���ͣ�ﰲ�ţ�", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!��ͣ��)) = 0 Then
                MsgBox "��ǰ�ϰ�ʱ����ͣ�ﰲ�ţ����ܵ�������������", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!��ʱ)) = 1 Then
                MsgBox "��ǰʱ���Ѵ�����ͣ�￪ʼʱ�䣬�����ٿ���ͣ�ﰲ�ţ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '"       And a.��ʼʱ�� <> a.��ֹʱ��" '��ʼʱ������ֹʱ����ȵ��ǼӺŵ����
            strSQL = "Select Sum(Decode(a.�Ƿ�ͣ��, 1, 0, 1) * Decode(Nvl(a.�Һ�״̬, 0), 0, 0, 1)) As ��С����," & vbNewLine & _
                    "        Count(1) As �������" & vbNewLine & _
                    " From �ٴ�������ſ��� A, �ٴ������¼ B" & vbNewLine & _
                    " Where a.��¼id = b.Id And b.Id = [1] And a.��ʼʱ�� Between b.ͣ�￪ʼʱ�� And b.ͣ����ֹʱ��" & vbNewLine & _
                    "       And a.��ʼʱ�� <> a.��ֹʱ��"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��鰲��", mlng��¼ID)
            If rsTemp.EOF Then
                MsgBox "δ�ҵ����ŵ�ͣ��ʱ�䷶Χ�ڵ����ʱ�����ݣ����ܿ���ͣ�ﰲ�ţ�", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!��С����)) > Val(txtOpen(0).Text) Then
                MsgBox "������������С����С��������" & Val(Nvl(rsTemp!��С����)) & "��", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!�������)) < Val(txtOpen(0).Text) Then
                MsgBox "������������С����С��������" & Val(Nvl(rsTemp!�������)) & "", vbInformation, gstrSysName
                Exit Sub
            End If
        
            'Procedure Zl_�ٴ�������ſ���_���ŹҺ�(
            strSQL = "Zl_�ٴ�������ſ���_���ŹҺ�("
            '  ��¼id_In �ٴ������¼.Id%Type,
            strSQL = strSQL & mlng��¼ID & ","
            '  ����_In Number
            strSQL = strSQL & Val(txtOpen(0).Text) & ")"
            
            gobjDatabase.ExecuteProcedure strSQL, Me.Caption
            mblnOK = True
            mblnChanged = False
            Unload Me
        End If
    Else
        For i = 0 To tbPage.ItemCount - 1
            If txtOpen(i).Enabled <> 0 Then
                strIDs = strIDs & "," & Val(tbPage(i).Tag)
            End If
        Next
        If strIDs = "" Then
            MsgBox "��ǰû�е����κΰ��ŵĿ������������豣�棡", vbInformation, gstrSysName
            Exit Sub
        Else
            strIDs = Mid(strIDs, 2)
        End If
        
        strSQL = "Select a.ID As ��¼ID, Nvl(a.�Ƿ���ſ���, 0) * Nvl(a.�Ƿ��ʱ��, 0) As �������ʱ��," & vbNewLine & _
                "        Decode(a.ͣ�￪ʼʱ��, Null, 0, 1) As ��ͣ��," & vbNewLine & _
                "        Decode(Sign(Nvl(a.ͣ�￪ʼʱ��, Sysdate) - Sysdate), -1, 1, 0) As ��ʱ" & vbNewLine & _
                " From �ٴ������¼ A, Table(f_Num2list([1])) B" & vbNewLine & _
                " Where a.Id = b.Column_Value"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��鰲��", strIDs)
        For i = 0 To tbPage.ItemCount - 1
            If txtOpen(i).Enabled Then
                rsTemp.Filter = "��¼ID=" & Val(tbPage(i).Tag)
                If rsTemp.EOF Then
                    MsgBox "δ�ҵ�[" & tbPage(i).Caption & "]�������ݣ����ܿ���ͣ�ﰲ�ţ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!�������ʱ��)) = 0 Then
                    MsgBox "[" & tbPage(i).Caption & "]���ﰲ�Ų������������������ʱ�εģ����ܿ���ͣ�ﰲ�ţ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!��ͣ��)) = 0 Then
                    MsgBox "[" & tbPage(i).Caption & "]ʱ����ͣ�ﰲ�ţ����ܵ�������������", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!��ʱ)) = 1 Then
                    MsgBox "��ǰʱ���Ѵ�����[" & tbPage(i).Caption & "]ͣ�￪ʼʱ�䣬�����ٿ���ͣ�ﰲ�ţ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '"       And a.��ʼʱ�� <> a.��ֹʱ��" '��ʼʱ������ֹʱ����ȵ��ǼӺŵ����
        strSQL = "Select a.��¼id, Sum(Decode(a.�Ƿ�ͣ��, 1, 0, 1) * Decode(Nvl(a.�Һ�״̬, 0), 0, 0, 1)) As ��С����," & vbNewLine & _
                "        Count(1) As �������" & vbNewLine & _
                " From �ٴ�������ſ��� A, �ٴ������¼ B, Table(f_Num2list([1])) C" & vbNewLine & _
                " Where a.��¼id = b.Id And b.Id = c.Column_Value" & vbNewLine & _
                "       And a.��ʼʱ�� Between b.ͣ�￪ʼʱ�� And b.ͣ����ֹʱ��" & vbNewLine & _
                "       And a.��ʼʱ�� <> a.��ֹʱ��" & vbNewLine & _
                " Group By a.��¼id"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��鰲�����", strIDs)
        For i = 0 To tbPage.ItemCount - 1
            If txtOpen(i).Enabled Then
                rsTemp.Filter = "��¼ID=" & Val(tbPage(i).Tag)
                If rsTemp.EOF Then
                    MsgBox "δ�ҵ�[" & tbPage(i).Caption & "]���ŵ�ͣ��ʱ�䷶Χ�ڵ����ʱ�����ݣ����ܿ���ͣ�ﰲ�ţ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!��С����)) > Val(txtOpen(0).Text) Then
                    MsgBox "[" & tbPage(i).Caption & "]�Ŀ�����������С����С��������" & Val(Nvl(rsTemp!��С����)) & "��", vbInformation, gstrSysName
                    tbPage(i).Selected = True
                    If txtOpen(i).Visible And txtOpen(i).Enabled Then txtOpen(i).SetFocus
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!�������)) < Val(txtOpen(0).Text) Then
                    MsgBox "[" & tbPage(i).Caption & "]�Ŀ�����������С����С��������" & Val(Nvl(rsTemp!�������)) & "", vbInformation, gstrSysName
                    tbPage(i).Selected = True
                    If txtOpen(i).Visible And txtOpen(i).Enabled Then txtOpen(i).SetFocus
                    Exit Sub
                End If
            End If
        Next
        
        Set cllSql = New Collection
        For i = 0 To tbPage.ItemCount - 1
            If txtOpen(i).Enabled Then
                'Procedure Zl_�ٴ�������ſ���_���ŹҺ�(
                strSQL = "Zl_�ٴ�������ſ���_���ŹҺ�("
                '  ��¼id_In �ٴ������¼.Id%Type,
                strSQL = strSQL & Val(tbPage(i).Tag) & ","
                '  ����_In Number
                strSQL = strSQL & Val(txtOpen(i).Text) & ")"
            
                zlAddArray cllSql, strSQL
            End If
        Next
        If cllSql.Count > 0 Then
            blnTrans = True
            zlExecuteProcedureArrAy cllSql, Me.Caption
            blnTrans = False
            mblnOK = True
            mblnChanged = False
            Unload Me
        End If
    End If
    Exit Sub
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        'Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub dtpDate_Change()
    Dim lngDept As Long, lngDoctor As Long
    
    On Error GoTo ErrHandler
    If cboDept.ListIndex <> -1 Then
        lngDept = cboDept.ItemData(cboDept.ListIndex)
    End If
    If cboDoctor.ListIndex <> -1 Then
        lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), lngDept, lngDoctor
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    On Error GoTo ErrHandler
    If cboDept.Visible And cboDept.Enabled Then
        cboDept.SetFocus
    ElseIf cboDoctor.Visible And cboDoctor.Enabled Then
        cboDoctor.SetFocus
    ElseIf dtpDate.Visible And dtpDate.Enabled Then
        dtpDate.SetFocus
    ElseIf txtOpen(0).Visible And txtOpen(0).Enabled Then
        txtOpen(0).SetFocus
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Function InitFace(ByVal lngModule As Long) As Boolean
    '��ʼ������
    Err = 0: On Error GoTo ErrHandler
    Select Case lngModule
    Case 1114 '�ٴ����ﰲ��
        picSignalSourceSelect.Visible = False
        vsfSignalSource.Visible = False
        fraSplit.Visible = False
        tbPage.Visible = False
        
        dtpDate.Enabled = False
        LoadControl 0
    Case Else
        fraSignalSource.Visible = False
        fraSplitY.Visible = False
        lblWorkTime.Visible = False: txtWorkTime.Visible = False
        Set fraRecordInfo(0).Container = picTimeWork(0)
        
        With tbPage
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
            
            .InsertItem 0, "���ϰ�ʱ��", picTimeWork(0).hWnd, 0
        End With
    End Select
    InitFace = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub Form_Load()
    mblnFirst = True
    If InitFace(mlngModule) = False Then Unload Me: Exit Sub
    If mlngModule = 1114 Then
        If LoadData(mlng��¼ID) = False Then Unload Me: Exit Sub
    Else
        cboDept.Enabled = mlngDeptID = 0
        cboDoctor.Enabled = mlngDoctorID = 0
        
        Call FillDept(mlngDeptID) '���������ҽ����ֻ��ѡ����Ա��������
        Call FillDoctor(, mlngDoctorID)
        dtpDate.Value = Format(gobjDatabase.CurrentDate(), "yyyy-MM-dd")
        dtpDate.minDate = dtpDate.Value
        Call dtpDate_Change
    End If
End Sub

Private Function LoadSignalSource(ByVal str�������� As String, _
    Optional ByVal lngDeptID As Long, _
    Optional ByVal lngDoctorID As Long) As Boolean
    '���غ�Դ����
    '��Σ�
    '   str�������� ��ʽ��yyyy-mm-dd
    '   lngDeptID ����ID
    '   lngDoctorID ҽ��ID
    Dim strSQL As String, strWhere As String
    Dim i As Integer, strIDs As String
    Dim lngRow As Long

    Err = 0: On Error GoTo ErrHandler
    vsfSignalSource.Clear 1: vsfSignalSource.Rows = 1
    For i = picTimeWork.UBound To 1 Step -1
        tbPage.RemoveItem i
        UnLoadControl i
    Next
    tbPage(0).Caption = "���ϰ�ʱ��"
    UnLoadControl 0
    mlngPreRow = -1
    
    Set mrsRecord = Nothing: Set mrsRecordCount = Nothing
    '��Դ��Ϣ
    If lngDeptID <> 0 Then strWhere = " And a.����ID=[1]"
    If lngDoctorID <> 0 Then strWhere = strWhere & " And a.ҽ��ID=[2]"
    strSQL = "Select a.��Դid, b.����, b.����, m.���� As ����, a.ҽ������ As ҽ��, n.���� As �շ���Ŀ," & vbNewLine & _
            "        a.Id As ��¼id, a.��������, a.�ϰ�ʱ��, a.�޺���, a.��Լ��, a.ԤԼ����," & vbNewLine & _
            "        a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��" & vbNewLine & _
            " From �ٴ������¼ A, �ٴ������Դ B, ���ű� M, �շ���ĿĿ¼ N" & vbNewLine & _
            " Where a.��Դid = b.Id And a.����id = m.Id And a.��Ŀid = n.Id And a.�������� = [3] And" & vbNewLine & _
            "       Nvl(a.�Ƿ���ſ���, 0) = 1 And Nvl(a.�Ƿ��ʱ��, 0) = 1" & strWhere & vbNewLine & _
            "       And (b.����ʱ�� Is Null Or b.����ʱ��>=To_Date('3000-01-01','yyyy-mm-dd'))" & vbNewLine & _
            "       And (m.վ��='" & gstrNodeNo & "' Or m.վ�� is Null)"
    Set mrsRecord = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Դ��Ϣ", lngDeptID, lngDoctorID, CDate(str��������))
    If mrsRecord.EOF Then Set mrsRecord = Nothing: Exit Function
    
    '"       And b.��ʼʱ�� <> b.��ֹʱ��" '��ʼʱ������ֹʱ����ȵ��ǼӺŵ����
    strSQL = "Select b.��¼ID, Nvl(Max(Decode(b.�Ƿ�ͣ��, 1, 0, b.���)) - Min(b.���) + 1, 0) As ͣ�ﷶΧ," & vbNewLine & _
            "        Sum(Decode(b.�Ƿ�ͣ��, 1, 0, 1) * Decode(Nvl(b.�Һ�״̬, 0), 0, 0, 1)) As ��С����," & vbNewLine & _
            "        Count(1) As �������," & vbNewLine & _
            "        Sum(Decode(b.�Ƿ�ͣ��, 1, 0, 1)) As �ϴο�������" & vbNewLine & _
            " From �ٴ�������ſ��� B, �ٴ������¼ A, �ٴ������Դ M, ���ű� N" & vbNewLine & _
            " Where b.��¼id = a.Id And a.��Դid = m.Id And a.����id = n.Id" & vbNewLine & _
            "       And b.��ʼʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ��" & vbNewLine & _
            "       And b.��ʼʱ�� <> b.��ֹʱ��" & vbNewLine & _
            "       And a.�������� = [3] And Nvl(a.�Ƿ���ſ���, 0) = 1 And Nvl(a.�Ƿ��ʱ��, 0) = 1" & strWhere & vbNewLine & _
            "       And (m.����ʱ�� Is Null Or m.����ʱ��>=To_Date('3000-01-01','yyyy-mm-dd'))" & vbNewLine & _
            "       And (n.վ��='" & gstrNodeNo & "' Or n.վ�� is Null)" & vbNewLine & _
            " Group By b.��¼ID"
    Set mrsRecordCount = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��ſ�������", lngDeptID, lngDoctorID, CDate(str��������))
    
    With vsfSignalSource
        .Redraw = flexRDNone
        lngRow = 1
        Do While Not mrsRecord.EOF
            If InStr("," & strIDs & ",", "," & Val(Nvl(mrsRecord!��ԴID)) & ",") = 0 Then
                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("��ԴID")) = Val(Nvl(mrsRecord!��ԴID))
                .TextMatrix(lngRow, .ColIndex("����")) = Nvl(mrsRecord!����)
                .TextMatrix(lngRow, .ColIndex("����")) = Nvl(mrsRecord!����)
                .TextMatrix(lngRow, .ColIndex("����")) = Nvl(mrsRecord!����)
                .TextMatrix(lngRow, .ColIndex("ҽ��")) = Nvl(mrsRecord!ҽ��)
                .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) = Nvl(mrsRecord!�շ���Ŀ)
                lngRow = lngRow + 1
            End If
            strIDs = strIDs & "," & Val(Nvl(mrsRecord!��ԴID))
            mrsRecord.MoveNext
        Loop
        .Redraw = flexRDBuffered
        If .Rows > 1 Then
            mblnNotChange = True
            .Row = 1
            mblnNotChange = False
            vsfSignalSource_EnterCell
        End If
    End With
    LoadSignalSource = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function LoadData(ByVal lng��¼ID As Long) As Boolean
    '��������
    Dim strSQL As String, rsRecord As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim lngCanOpenMax As Long, lngCanOpenMin As Long, lngPreOpen As Long

    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select b.����, b.����, c.���� As ����, a.ҽ������, d.���� As �շ���Ŀ," & vbNewLine & _
            "        a.Id As ��¼id, a.��������, a.�ϰ�ʱ��, a.�޺���, a.��Լ��, a.ԤԼ����," & vbNewLine & _
            "        a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��" & vbNewLine & _
            " From �ٴ������¼ A, �ٴ������Դ B, ���ű� C, �շ���ĿĿ¼ D" & vbNewLine & _
            " Where a.��ԴID = b.ID And a.����id = c.Id And a.��Ŀid = d.Id" & vbNewLine & _
            "       And a.id = [1] And Nvl(a.�Ƿ���ſ���, 0) = 1 And Nvl(a.�Ƿ��ʱ��, 0) = 1"
    Set rsRecord = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ����", lng��¼ID)
    If rsRecord.EOF Then
        MsgBox "��ǰ����δ������ſ��ƻ�δ���÷�ʱ�Σ����ܿ���ͣ�ﰲ�ţ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    txtSignalNO.Text = Nvl(rsRecord!����)
    txt����.Text = Nvl(rsRecord!����)
    txtDept.Text = Nvl(rsRecord!����)
    txtDoctor.Text = Nvl(rsRecord!ҽ������)
    txtItem.Text = Nvl(rsRecord!�շ���Ŀ)
    
    dtpDate.Value = Format(Nvl(rsRecord!��������), "yyyy-MM-dd")
    txtWorkTime.Text = Nvl(rsRecord!�ϰ�ʱ��)
    txt�޺���(0).Text = IIf(Val(Nvl(rsRecord!�޺���)) = 0, "", Nvl(rsRecord!�޺���))
    txt��Լ��(0).Text = IIf(Val(Nvl(rsRecord!ԤԼ����)) = 1, "��ֹԤԼ", IIf(Val(Nvl(rsRecord!��Լ��)) = 0, txt�޺���(0).Text, Nvl(rsRecord!��Լ��)))
    txtͣ��ʱ��(0).Text = Format(Nvl(rsRecord!ͣ�￪ʼʱ��), "hh:mm") & "��" & Format(Nvl(rsRecord!ͣ����ֹʱ��), "hh:mm")
    txtͣ��ʱ��(0).Tag = Format(Nvl(rsRecord!ͣ�￪ʼʱ��), "yyyy-mm-dd hh:mm") & "��" & Format(Nvl(rsRecord!ͣ����ֹʱ��), "yyyy-mm-dd hh:mm")
    If txtͣ��ʱ��(0).Text = "��" Then txtͣ��ʱ��(0).Text = ""
    
    '"       And a.��ʼʱ�� <> a.��ֹʱ��" '��ʼʱ������ֹʱ����ȵ��ǼӺŵ����
    strSQL = "Select Nvl(Max(Decode(a.�Ƿ�ͣ��, 1, 0, a.���)) - Min(a.���) + 1, 0) As ͣ�ﷶΧ," & vbNewLine & _
            "        Sum(Decode(a.�Ƿ�ͣ��, 1, 0, 1) * Decode(Nvl(a.�Һ�״̬, 0), 0, 0, 1)) As ��С����," & vbNewLine & _
            "        Count(1) As �������," & vbNewLine & _
            "        Sum(Decode(a.�Ƿ�ͣ��, 1, 0, 1)) As �ϴο�������" & vbNewLine & _
            " From �ٴ�������ſ��� A, �ٴ������¼ B" & vbNewLine & _
            " Where a.��¼id = b.Id And b.Id = [1]" & vbNewLine & _
            "       And a.��ʼʱ�� Between b.ͣ�￪ʼʱ�� And b.ͣ����ֹʱ��" & vbNewLine & _
            "       And a.��ʼʱ�� <> a.��ֹʱ��"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡͣ���������", lng��¼ID)
    If Not rsTemp.EOF Then
        vsfTimeWork(0).Tag = Val(Nvl(rsTemp!ͣ�ﷶΧ))
        lngCanOpenMax = Val(Nvl(rsTemp!�������))
        lngCanOpenMin = Val(Nvl(rsTemp!��С����))
        lngPreOpen = Val(Nvl(rsTemp!�ϴο�������))
    End If
    mblnNotChange = True
    updOpen(0).Max = lngCanOpenMax
    updOpen(0).Min = lngCanOpenMin
    txtOpen(0).Text = lngPreOpen
    mblnNotChange = False
    
    '���ʱ��
    Call LoadDataToGrid(0, Val(Nvl(rsRecord!��¼ID)), lngCanOpenMax, lngCanOpenMin)
    mblnChanged = False
    LoadData = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function LoadDataToGrid(ByVal Index As Integer, ByVal lng��¼ID As Long, _
    ByVal lngCanOpenMax As Long, ByVal lngCanOpenMin As Long) As Boolean
    '����ʱ�����ݵ�����ؼ�
    '���:
    '   lng��¼ID - ��¼ID
    '����:���سɹ�������true,���򷵻�Flase
    Dim objColAll As New Collection 'Array(���,��ʼʱ��,��ֹʱ��,�Һ�״̬,�Ƿ�ͣ��)
    Dim strSQL As String, rsRecord As ADODB.Recordset

    On Error GoTo ErrHander
    '�������ʱ���Ⱥ�������򣬲�ȻҪ��
    '"       And a.��ʼʱ�� <> a.��ֹʱ��" '��ʼʱ������ֹʱ����ȵ��ǼӺŵ����
    strSQL = "Select a.���, a.��ʼʱ��, a. ��ֹʱ��, a.�Һ�״̬, a.�Ƿ�ͣ��" & vbNewLine & _
            " From �ٴ�������ſ��� A" & vbNewLine & _
            " Where a.��¼ID = [1] And a.��ʼʱ�� <> a.��ֹʱ��" & vbNewLine & _
            " Order By a.���"
    Set rsRecord = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng��¼ID)
    Do While Not rsRecord.EOF
        objColAll.Add Array(Val(Nvl(rsRecord!���)), Format(Nvl(rsRecord!��ʼʱ��), "yyyy-MM-dd hh:mm:ss"), _
            Format(Nvl(rsRecord!��ֹʱ��), "yyyy-MM-dd hh:mm:ss"), _
            Val(Nvl(rsRecord!�Һ�״̬)), Val(Nvl(rsRecord!�Ƿ�ͣ��)))
        rsRecord.MoveNext
    Loop
    LoadDataToGrid = ShowTimeIntervals(Index, objColAll, lngCanOpenMax, lngCanOpenMin)
    Exit Function
ErrHander:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function ShowTimeIntervals(ByVal Index As Integer, objCol As Collection, _
    ByVal lngCanOpenMax As Long, ByVal lngCanOpenMin As Long) As Boolean
    '��ʾʱ������
    '��Σ�
    '   objCol:Array(���,��ʼʱ��,��ֹʱ��,�Һ�״̬,�Ƿ�ͣ��)
    Dim varItem As Variant, varTemp As Variant
    Dim i As Integer, j As Integer, blnFind As Boolean
    Dim lngRow As Long, lngCol As Long, strCurTime As String
    Dim dtSys As Date, strToolTip As String
    Dim strStopStart As String, strStopEnd As String

    Err = 0: On Error GoTo ErrHander:
    lblToolTip(Index).Caption = ""
    With vsfTimeWork(Index)
        .Clear
        .Rows = 0
        If objCol Is Nothing Then Exit Function
        If objCol.Count = 0 Then Exit Function
        
        .Redraw = flexRDNone
        dtSys = gobjDatabase.CurrentDate
        strStopStart = Split(txtͣ��ʱ��(Index).Tag & "��", "��")(0)
        strStopEnd = Split(txtͣ��ʱ��(Index).Tag & "��", "��")(1)
        If IsDate(strStopStart) Then
            If DateDiff("n", strStopStart, dtSys) > 0 Then
               '��ǰʱ���ѽ���ͣ��ʱ�䷶Χ
               strToolTip = "��ǰʱ���Ѵ���ͣ�￪ʼʱ�䣬���ܵ�������������"
            End If
        Else
            '��ǰ�ϰ�ʱ����ͣ�ﰲ��
            strToolTip = "��ǰ�ϰ�ʱ����ͣ�ﰲ�ţ����ܵ�������������"
        End If
        
        .Rows = 1: .Cols = 1
        .FixedCols = 1

        lngRow = -1: lngCol = 1: strCurTime = ""
        .FontSize = 9
        For Each varItem In objCol
            If strCurTime <> Format(varItem(1), "hh:00") Then
                strCurTime = Format(varItem(1), "hh:00")
                lngRow = lngRow + 1: lngCol = 1
                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(lngRow, 0) = strCurTime
            End If
            If lngCol > .Cols - 1 Then .Cols = .Cols + 1
            .TextMatrix(lngRow, lngCol) = varItem(0) & vbCrLf & _
                Format(varItem(1), "hh:mm") & "-" & Format(varItem(2), "hh:mm")
            .Cell(flexcpData, lngRow, lngCol) = Format(varItem(1), "yyyy-MM-dd hh:mm:ss") & "��" & Format(varItem(2), "yyyy-MM-dd hh:mm:ss")
            
            If Format(dtSys, "yyyy-mm-dd hh:mm:ss") >= Format(varItem(1), "yyyy-mm-dd hh:mm:ss") Then
                 '��ʧЧ�����»��ߺͻ�ɫ������ʾ
                 .Cell(flexcpFontUnderline, lngRow, lngCol) = True
                 .Cell(flexcpForeColor, lngRow, lngCol) = vbGrayText
            End If
            
            Select Case varItem(3)
            Case SNState.�ѹ�
                .Cell(flexcpForeColor, lngRow, lngCol) = &HC0&
                .Cell(flexcpFontStrikethru, lngRow, lngCol) = True
            Case SNState.��Լ
                .Cell(flexcpForeColor, lngRow, lngCol) = vbGreen
            Case SNState.Ԥ��
                .Cell(flexcpForeColor, lngRow, lngCol) = vbBlue
            Case SNState.�˺�
                .Cell(flexcpForeColor, lngRow, lngCol) = vbGrayText
                .Cell(flexcpFontStrikethru, lngRow, lngCol) = True
            Case SNState.����
                .Cell(flexcpForeColor, lngRow, lngCol) = &HC0&
            End Select
            
            '�ж��Ƿ���ͣ�ﷶΧ�ڿ��ŵ�
            If varItem(4) = 1 Then
                '��ͣ���ú�ɫ������ʾ
                .Cell(flexcpBackColor, lngRow, lngCol) = vbRed
            End If
            
            lngCol = lngCol + 1
        Next
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = flexAlignCenterTop
        .Cell(flexcpAlignment, 0, 1, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpFontSize, 0, 0, .Rows - 1) = 12
        .Cell(flexcpFontBold, 0, 0, .Rows - 1) = True
        .ColWidth(-1) = 1100: .ColWidth(0) = 1000: .RowHeight(-1) = 600
        
        If strToolTip = "" Then
            If lngCanOpenMax = 0 Then
                strToolTip = "ͣ��ʱ�䷶Χ�ڲ����ڡ����һ��ԤԼ�ĺš������ܵ�������������"
                txtOpen(Index).Enabled = False: updOpen(Index).Enabled = False
            ElseIf lngCanOpenMin > 0 Then
                strToolTip = "�ϴο��ŵ����������� " & lngCanOpenMin & " ����ʹ�ã��������õĿ�������������� " & lngCanOpenMin & " ��"
            End If
        Else
            txtOpen(Index).Enabled = False: updOpen(Index).Enabled = False
        End If
        If Trim(strToolTip) = "" Then
            lblToolTip(Index).Caption = ""
        Else
            lblToolTip(Index).Caption = "��ʾ��" & strToolTip
        End If
        .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftTop
        .Redraw = flexRDBuffered
    End With
    ShowTimeIntervals = True
    Exit Function
ErrHander:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Select Case mlngModule
    Case 1114 '�ٴ����ﰲ��
        fraSignalSource.Move 0, 10, Me.ScaleWidth
        picVisitDate.Move 0, fraSignalSource.Top + fraSignalSource.Height
        fraRecordInfo(0).Move picVisitDate.Left + picVisitDate.Width, picVisitDate.Top
        With picTimeWork(0)
            .Left = 0
            .Top = picVisitDate.Top + picVisitDate.Height
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight - 800 - .Top
        End With
        fraSplitY.Move -10, picTimeWork(0).Top + picTimeWork(0).Height, Me.ScaleWidth + 20
    Case Else
        picSignalSourceSelect.Move 0, 50
        picVisitDate.Move picSignalSourceSelect.Left + picSignalSourceSelect.Width, picSignalSourceSelect.Top
        
        With vsfSignalSource
            .Left = 0
            .Top = picSignalSourceSelect.Top + picSignalSourceSelect.Height
            .Height = Me.ScaleHeight - 800 - .Top
        End With
        fraSplit.Move vsfSignalSource.Left + vsfSignalSource.Width, vsfSignalSource.Top, fraSplit.Width, Me.ScaleHeight - vsfSignalSource.Top - 800
        With tbPage
            .Left = fraSplit.Left + fraSplit.Width
            .Top = fraSplit.Top
            .Width = Me.ScaleWidth - .Left
            .Height = fraSplit.Height
        End With
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If mblnChanged Then
        If MsgBox("��ǰ���ŵĿ��������Ѹı䣬������δ���棬�Ƿ񲻱��棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    Set mrsDept = Nothing
    Set mrsDoctor = Nothing
    Set mrsRecord = Nothing: Set mrsRecordCount = Nothing
End Sub

Private Sub fraSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then msngStartX = X
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    
    On Error Resume Next
    If Button = vbLeftButton Then
        sngTemp = fraSplit.Left + X - msngStartX
        If sngTemp > 500 And Me.ScaleWidth - (sngTemp + fraSplit.Width) > 500 Then
            fraSplit.Left = sngTemp
            vsfSignalSource.Width = fraSplit.Left - vsfSignalSource.Left
            tbPage.Move fraSplit.Left + fraSplit.Width, tbPage.Top, Me.ScaleWidth - (fraSplit.Left + fraSplit.Width)
        End If
    End If
End Sub

Private Sub picTimeWork_Resize(Index As Integer)
    Dim blnRecordInfo As Boolean
    
    On Error Resume Next
    If Index = 0 And fraRecordInfo(0).Container <> picTimeWork(0) Then
        vsfTimeWork(Index).Move 0, 30, picTimeWork(Index).ScaleWidth, picTimeWork(Index).ScaleHeight - 450 - 60
    Else
        fraRecordInfo(Index).Move 0, 0
        vsfTimeWork(Index).Move 0, 420, picTimeWork(Index).ScaleWidth, picTimeWork(Index).ScaleHeight - 450 - 420
    End If
    txtOpen(Index).Top = picTimeWork(Index).ScaleHeight - txtOpen(Index).Height - 60
    updOpen(Index).Top = txtOpen(Index).Top
    lblOpen(Index).Top = txtOpen(Index).Top + (txtOpen(Index).Height - lblOpen(Index).Height) / 2
    lblToolTip(Index).Top = lblOpen(Index).Top
    lblToolTip(Index).Width = picTimeWork(Index).ScaleWidth - lblToolTip(Index).Left
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If txtOpen(Item.Index).Visible And txtOpen(Item.Index).Enabled Then txtOpen(Item.Index).SetFocus
End Sub

Private Sub txtOpen_Change(Index As Integer)
    If mblnNotChange Then Exit Sub
    
    On Error GoTo ErrHandler
    mblnChanged = True
    mblnNotChange = True
    If Trim(txtOpen(Index).Text) = "" Then txtOpen(Index).Text = "0"
    If updOpen(Index).Max < Val(txtOpen(Index).Text) Then
        MsgBox "�����������ܴ������ɿ�������(" & updOpen(Index).Max & ")��", vbExclamation, gstrSysName
        txtOpen(Index).Text = updOpen(Index).Max
        If txtOpen(Index).Visible And txtOpen(Index).Enabled Then txtOpen(Index).SetFocus
    End If
    If updOpen(Index).Min > Val(txtOpen(Index).Text) Then
        MsgBox "������������С����С�ɿ�������(" & updOpen(Index).Min & ")��", vbExclamation, gstrSysName
        txtOpen(Index).Text = updOpen(Index).Min
        If txtOpen(Index).Visible And txtOpen(Index).Enabled Then txtOpen(Index).SetFocus
    End If
    mblnNotChange = False
    Call OpenSN(Index, Val(txtOpen(Index).Text))
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub OpenSN(Index As Integer, ByVal lngCount As Long)
    '�������
    Dim lngRow As Long, lngCol As Long
    Dim strStopStart As String, strStopEnd As String
    Dim strStart As String, blnStart As Boolean
    
    On Error GoTo ErrHandler
    strStopStart = Split(txtͣ��ʱ��(Index).Tag & "��", "��")(0)
    strStopEnd = Split(txtͣ��ʱ��(Index).Tag & "��", "��")(1)
    If Not (IsDate(strStopStart) And IsDate(strStopEnd)) Then Exit Sub
    With vsfTimeWork(Index)
        .Redraw = flexRDNone
        
        If Val(vsfTimeWork(Index).Tag) > lngCount Then
            OpenSN Index, vsfTimeWork(Index).Tag '�Ȱ����ӻָ���ʼ״̬����֤��ͣ�￪ʼ��ŵ����ŵ������Ŷ���������
            lngCount = Val(vsfTimeWork(Index).Tag) - lngCount
            
            '���ٿ�������
            For lngRow = .Rows - 1 To .FixedRows Step -1
                For lngCol = .Cols - 1 To .FixedCols Step -1
                    If .TextMatrix(lngRow, lngCol) <> "" Then
                        strStart = Split(.Cell(flexcpData, lngRow, lngCol), "��")(0)
                        If DateDiff("n", strStopStart, strStart) >= 0 And DateDiff("n", strStopEnd, strStart) <= 0 Then
                            If .Cell(flexcpBackColor, lngRow, lngCol) = .BackColor And blnStart = False Then
                                blnStart = True '��ǿ�ʼ
                            End If
                            If blnStart And lngCount > 0 Then
                                If (.Cell(flexcpForeColor, lngRow, lngCol)) = vbBlack Then '��ɫ����ĹҺ�״̬Ϊ"0-���һ��ԤԼ�ĺ�"
                                    .Cell(flexcpBackColor, lngRow, lngCol) = vbRed
                                    lngCount = lngCount - 1
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        Else
            For lngRow = .FixedRows To .Rows - 1
                For lngCol = .FixedCols To .Cols - 1
                    If .TextMatrix(lngRow, lngCol) <> "" Then
                        strStart = Split(.Cell(flexcpData, lngRow, lngCol), "��")(0)
                        If DateDiff("n", strStopStart, strStart) >= 0 And DateDiff("n", strStopEnd, strStart) <= 0 Then
                            If lngCount > 0 Then
                                .Cell(flexcpBackColor, lngRow, lngCol) = .BackColor
                                lngCount = lngCount - 1
                            Else
                                .Cell(flexcpBackColor, lngRow, lngCol) = vbRed
                            End If
                        End If
                    End If
                Next
            Next
        End If
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub txtOpen_GotFocus(Index As Integer)
    gobjControl.TxtSelAll txtOpen(Index)
End Sub

Private Sub txtOpen_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub LoadControl(ByVal Index As Long)
    '���ӿؼ�
    On Error GoTo ErrHandler
    If ExistsControl(picTimeWork(Index)) Then
        txt�޺���(Index).Text = ""
        txt��Լ��(Index).Text = ""
        txtͣ��ʱ��(Index).Text = "": txtͣ��ʱ��(Index).Tag = ""
        vsfTimeWork(Index).Clear: vsfTimeWork(Index).Rows = 0
        mblnNotChange = True
        txtOpen(Index).Text = "0"
        mblnNotChange = False
    Else
        Load picTimeWork(Index): picTimeWork(Index).Visible = True
        
        Load lbl�޺���(Index): lbl�޺���(Index).Visible = True: Set lbl�޺���(Index).Container = picTimeWork(Index)
        Load txt�޺���(Index): txt�޺���(Index).Visible = True: Set txt�޺���(Index).Container = picTimeWork(Index)
        Load lbl��Լ��(Index): lbl��Լ��(Index).Visible = True: Set lbl��Լ��(Index).Container = picTimeWork(Index)
        Load txt��Լ��(Index): txt��Լ��(Index).Visible = True: Set txt��Լ��(Index).Container = picTimeWork(Index)
        Load lblͣ��ʱ��(Index): lblͣ��ʱ��(Index).Visible = True: Set lblͣ��ʱ��(Index).Container = picTimeWork(Index)
        Load txtͣ��ʱ��(Index): txtͣ��ʱ��(Index).Visible = True: Set txtͣ��ʱ��(Index).Container = picTimeWork(Index)
        
        Load vsfTimeWork(Index): vsfTimeWork(Index).Visible = True: Set vsfTimeWork(Index).Container = picTimeWork(Index)
        
        Load lblOpen(Index): lblOpen(Index).Visible = True: Set lblOpen(Index).Container = picTimeWork(Index)
        Load txtOpen(Index): txtOpen(Index).Visible = True: Set txtOpen(Index).Container = picTimeWork(Index)
        Load updOpen(Index): updOpen(Index).Visible = True: Set updOpen(Index).Container = picTimeWork(Index)
        updOpen(Index).BuddyControl = txtOpen(Index): updOpen(Index).BuddyProperty = "Text"
        Load lblToolTip(Index): lblToolTip(Index).Visible = True: Set lblToolTip(Index).Container = picTimeWork(Index)
    End If
    txtOpen(Index).Enabled = True: updOpen(Index).Enabled = True
    
    '���⴦��һ�£���Ϊ�ڶ�̬����updOpenʱǰһ��txtOpen�Ŀ�Ȼ��
    Dim i As Integer
    For i = txtOpen.LBound To txtOpen.UBound
        txtOpen(i).Width = 1100: updOpen(i).Left = txtOpen(i).Left + txtOpen(i).Width + 10
    Next
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub UnLoadControl(ByVal Index As Long)
    'ж�ؿؼ�
    On Error GoTo ErrHandler
    If ExistsControl(picTimeWork(Index)) Then
        If Index = 0 Then
            txt�޺���(Index).Text = ""
            txt��Լ��(Index).Text = ""
            txtͣ��ʱ��(Index).Text = "": txtͣ��ʱ��(Index).Tag = ""
            vsfTimeWork(Index).Clear: vsfTimeWork(Index).Rows = 0
            mblnNotChange = True
            txtOpen(Index).Text = "0"
            mblnNotChange = False
            txtOpen(Index).Enabled = False: updOpen(Index).Enabled = False
            lblToolTip(Index).Caption = ""
            tbPage(Index).Tag = ""
        Else
            '����ж�أ���ComboBox_Click�в���ж�ؿؼ�������"�����ڸ���������ж�أ����� 365��"
'            Unload lbl�޺���(index): Unload txt�޺���(index)
'            Unload lbl��Լ��(index): Unload txt��Լ��(index)
'            Unload lblͣ��ʱ��(index): Unload txtͣ��ʱ��(index)
'
'            Unload vsfTimeWork(index)
'
'            Unload lblOpen(index): Unload txtOpen(index): Unload updOpen(index): Unload lblToolTip(index)
'
'            Unload picTimeWork(index)
        End If
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Function ExistsControl(ByRef ctlVal As Control) As Boolean
    '�жϿؼ��Ƿ�ʵ��
    Dim strTmp As String

    On Error GoTo ErrHandler
    strTmp = ctlVal.Name
    ExistsControl = True
    Exit Function
ErrHandler:
    ExistsControl = False
End Function

Private Sub vsfSignalSource_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    mlngPreRow = NewRow
End Sub

Private Sub vsfSignalSource_EnterCell()
    Dim lng��ԴID As Long
    Dim lngRow As Long, i As Integer
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngCanOpenMax As Long, lngCanOpenMin As Long, lngPreOpen As Long

    Err = 0: On Error GoTo ErrHandler
    If mblnNotChange Then Exit Sub
    If vsfSignalSource.Row < vsfSignalSource.FixedRows Then Exit Sub
    If mrsRecord Is Nothing Then Exit Sub
    
    If mblnChanged Then
        If MsgBox("��ǰ���ŵĿ��������Ѹı䣬������δ���棬�Ƿ񲻱��棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            mblnNotChange = True
            vsfSignalSource.Row = mlngPreRow
            mblnNotChange = False
            Exit Sub
        Else
            mblnChanged = False
        End If
    End If
    
    For i = picTimeWork.UBound To 1 Step -1
        tbPage.RemoveItem i
        UnLoadControl i
    Next
    tbPage(0).Caption = "���ϰ�ʱ��"
    UnLoadControl 0
    
    lng��ԴID = Val(vsfSignalSource.TextMatrix(vsfSignalSource.Row, vsfSignalSource.ColIndex("��ԴID")))
    mrsRecord.Filter = "��ԴID=" & lng��ԴID
    lngRow = 0
    Do While Not mrsRecord.EOF
        LoadControl lngRow
        
        txt�޺���(lngRow).Text = IIf(Val(Nvl(mrsRecord!�޺���)) = 0, "", Nvl(mrsRecord!�޺���))
        txt��Լ��(lngRow).Text = IIf(Val(Nvl(mrsRecord!ԤԼ����)) = 1, "��ֹԤԼ", IIf(Val(Nvl(mrsRecord!��Լ��)) = 0, txt�޺���(lngRow).Text, Nvl(mrsRecord!��Լ��)))
        txtͣ��ʱ��(lngRow).Text = Format(Nvl(mrsRecord!ͣ�￪ʼʱ��), "hh:mm") & "��" & Format(Nvl(mrsRecord!ͣ����ֹʱ��), "hh:mm")
        txtͣ��ʱ��(lngRow).Tag = Format(Nvl(mrsRecord!ͣ�￪ʼʱ��), "yyyy-mm-dd hh:mm") & "��" & Format(Nvl(mrsRecord!ͣ����ֹʱ��), "yyyy-mm-dd hh:mm")
        If txtͣ��ʱ��(lngRow).Text = "��" Then txtͣ��ʱ��(0).Text = ""
        
        If Not mrsRecordCount Is Nothing Then
            mrsRecordCount.Filter = "��¼ID=" & Val(Nvl(mrsRecord!��¼ID))
            If Not mrsRecordCount.EOF Then
                vsfTimeWork(lngRow).Tag = Val(Nvl(mrsRecordCount!ͣ�ﷶΧ))
                lngCanOpenMax = Val(Nvl(mrsRecordCount!�������))
                lngCanOpenMin = Val(Nvl(mrsRecordCount!��С����))
                lngPreOpen = Val(Nvl(mrsRecordCount!�ϴο�������))
            End If
        End If
        
        mblnNotChange = True
        updOpen(lngRow).Max = lngCanOpenMax
        updOpen(lngRow).Min = lngCanOpenMin
        txtOpen(lngRow).Text = lngPreOpen
        mblnNotChange = False
        
        LoadDataToGrid lngRow, Val(Nvl(mrsRecord!��¼ID)), lngCanOpenMax, lngCanOpenMin
        
        If lngRow = 0 Then
            tbPage(0).Caption = Nvl(mrsRecord!�ϰ�ʱ��)
        Else
            tbPage.InsertItem lngRow, Nvl(mrsRecord!�ϰ�ʱ��), picTimeWork(lngRow).hWnd, 0
        End If
        tbPage(lngRow).Tag = Nvl(mrsRecord!��¼ID)
        
        lngRow = lngRow + 1
        mrsRecord.MoveNext
    Loop
    If txtOpen(0).Visible And txtOpen(0).Enabled Then txtOpen(0).SetFocus
    mblnChanged = False
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub vsfTimeWork_BeforeRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    If vsfTimeWork(Index).TextMatrix(NewRow, NewCol) = "" Then Cancel = True: Exit Sub
End Sub

Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = True, Optional str���в��� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���:cboDept-ָ���Ĳ��Ų���
    '     rsDept-ָ���Ĳ���
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str���в���-���в�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String
    
    On Error GoTo ErrHandler
    '�ȸ��Ƽ�¼��
    Set rsTemp = gobjDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(gSysPara.strLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf gobjCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���в��� <> "" Then
        str���� = gobjCommFun.SpellCode(str���в���)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str���в���) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!����) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!����)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    End Select
    
    '����ѡ����
    If gobjDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "" & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then gobjCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.TxtSelAll cboDept
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPersonSelect(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboSel As ComboBox, ByVal rsPerson As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = True, Optional str���� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Աѡ��ѡ����
    '���:cboSel-ָ���Ĳ���ѡ�񲿼�
    '     rsPerson-ָ������Ա��Ϣ(ID,���,����,����)
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str����-��������(������,���в���Ա��)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String, strLike As String
    
    On Error GoTo ErrHandler
    '�ȸ��Ƽ�¼��
    Set rsTemp = gobjDatabase.zlCopyDataStructure(rsPerson)
    
    strSearch = UCase(strSearch)
        
    strCompents = Replace(gSysPara.strLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf gobjCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���� <> "" Then
        str���� = gobjCommFun.SpellCode(str����)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    strIDs = ","
    With rsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!���) = strSearch Then lngID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!���)) = Val(strSearch) Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!���) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!���) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!���)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngID = 0
    If lngID <> 0 And rsTemp.RecordCount = 1 Then lngID = Nvl(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngID <> 0 Then GoTo GoOver:
    If lngID < 0 Then lngID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    End Select
    
    '����ѡ����
    If gobjDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboSel, rsTemp, True, "", "����ID" & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngID = Val(Nvl(rsReturn!ID))
    If lngID < 0 Then lngID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.CboLocate cboSel, lngID, True
    If blnSendKeys Then gobjCommFun.PressKey vbKeyTab
zlPersonSelect = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.TxtSelAll cboSel
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAllDoctor() As ADODB.Recordset
    '��ȡҽ���б�
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = "Select c.id,c.���,c.����,c.����,b.����id" & vbNewLine & _
            " From ��Ա����˵�� A, ������Ա B, ��Ա�� C" & vbNewLine & _
            " Where b.��Աid=c.id And b.��Աid=a.��Աid And a.��Ա����=[1]" & vbNewLine & _
            "       And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
            "       And (c.վ��=[2] Or c.վ�� is Null)" & vbNewLine & _
            " Order by c.���"
    Set GetAllDoctor = gobjDatabase.OpenSQLRecord(strSQL, "��ȡҽ��", "ҽ��", gstrNodeNo)
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetDepartments(ByVal str���� As String, _
    ByVal str������� As String, _
    Optional ByVal lng��Աid As Long = 0, _
    Optional ByVal blnCheckվ�� As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ʵĲ����б�
    '���:str����='�ٴ�','����','��ҩ��',...,����Ϊ��
    '     str�������:��,����:��1,3
    '     lng��ԱID-������0������Ա����������
    '����:
    '����:
    '����:���˺�
    '����:2009-10-12 09:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    str���� = Replace(str����, "'", "")
    If str���� <> "" Then
        If InStr(1, str����, ",") > 0 Then
            strSQL = " And Instr(','||[1]||',',','||B.��������||',')>0"
        Else
            strSQL = " And B.�������� = [1]"
        End If
    End If
    If lng��Աid <> 0 Then strSQL = strSQL & "  And A.id=C.����ID and C.��Աid =[3]"
    
    strSQL = _
        " Select Distinct A.ID,A.����,A.����,A.���� " & _
        " From ���ű� A,��������˵�� B " & IIf(lng��Աid <> 0, ",������Ա C", "") & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And Instr(',' || [2]|| ',',',' || B.������� || ',')>0 " & strSQL & _
         IIf(blnCheckվ��, " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
        " Order by A.����"
    Set GetDepartments = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ����", str����, str�������, lng��Աid)
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

