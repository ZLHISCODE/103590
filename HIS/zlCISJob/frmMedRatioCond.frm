VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMedRatioCond 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picCond 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   3405
      ScaleHeight     =   4290
      ScaleWidth      =   2970
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   90
      Width           =   2970
      Begin VB.CheckBox chkDrug 
         BackColor       =   &H80000005&
         Caption         =   "����ͳ��ʱҩƷ������ҩ����ҩ����ҩ��"
         Height          =   420
         Left            =   75
         TabIndex        =   30
         Top             =   3360
         Width           =   2640
      End
      Begin VB.Frame fraWay 
         BackColor       =   &H80000005&
         Caption         =   "ͳ�Ʒ�ʽ"
         Height          =   1455
         Left            =   30
         TabIndex        =   20
         Top             =   1725
         Width           =   2670
         Begin VB.PictureBox picPatiType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   390
            ScaleHeight     =   270
            ScaleWidth      =   2025
            TabIndex        =   31
            Top             =   1125
            Width           =   2025
            Begin VB.OptionButton optPatiType 
               BackColor       =   &H80000005&
               Caption         =   "��Ժ"
               Height          =   255
               Index           =   1
               Left            =   1275
               TabIndex        =   34
               Top             =   -15
               Width           =   660
            End
            Begin VB.OptionButton optPatiType 
               BackColor       =   &H80000005&
               Caption         =   "��Ժ"
               Height          =   255
               Index           =   0
               Left            =   450
               TabIndex        =   33
               Top             =   -15
               Value           =   -1  'True
               Width           =   750
            End
            Begin VB.Label lblPatiType 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Width           =   360
            End
         End
         Begin VB.OptionButton optWay 
            BackColor       =   &H80000005&
            Caption         =   "������"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   870
            Width           =   1410
         End
         Begin VB.OptionButton optWay 
            BackColor       =   &H80000005&
            Caption         =   "��������"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   570
            Width           =   1650
         End
         Begin VB.OptionButton optWay 
            BackColor       =   &H80000005&
            Caption         =   "����������"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   255
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame fraRange 
         BackColor       =   &H80000005&
         Caption         =   "���÷�Χ"
         Height          =   580
         Left            =   45
         TabIndex        =   16
         Top             =   0
         Width           =   2670
         Begin VB.OptionButton optRan 
            BackColor       =   &H80000005&
            Caption         =   "ȫԺ"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   19
            Top             =   270
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton optRan 
            BackColor       =   &H80000005&
            Caption         =   "סԺ"
            Height          =   195
            Index           =   2
            Left            =   930
            TabIndex        =   18
            Top             =   270
            Width           =   705
         End
         Begin VB.OptionButton optRan 
            BackColor       =   &H80000005&
            Caption         =   "����"
            Height          =   195
            Index           =   1
            Left            =   1740
            TabIndex        =   17
            Top             =   270
            Width           =   705
         End
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "��ѯ(&O)"
         Height          =   330
         Left            =   1800
         TabIndex        =   15
         Top             =   3840
         Width           =   870
      End
      Begin VB.ComboBox cboTim 
         Height          =   300
         Left            =   885
         TabIndex        =   14
         Text            =   "Combo2"
         Top             =   660
         Width           =   1800
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   1
         Left            =   885
         TabIndex        =   24
         Top             =   1320
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   106168323
         CurrentDate     =   41636
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   0
         Left            =   885
         TabIndex        =   25
         Top             =   1005
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   106168323
         CurrentDate     =   37952
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   195
         Index           =   1
         Left            =   645
         TabIndex        =   28
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label lblBegin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   195
         Index           =   0
         Left            =   645
         TabIndex        =   27
         Top             =   1065
         Width           =   180
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ�䷶Χ"
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   690
         Width           =   720
      End
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   6510
      ScaleHeight     =   4065
      ScaleWidth      =   2850
      TabIndex        =   0
      Top             =   930
      Width           =   2850
      Begin VB.Frame fraOutTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   45
         TabIndex        =   35
         Top             =   735
         Width           =   2760
         Begin VB.ComboBox cboOutTime 
            Height          =   300
            Left            =   795
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   15
            Width           =   1305
         End
         Begin VB.Label lblOutTime 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��Ժʱ��"
            Height          =   180
            Left            =   0
            TabIndex        =   36
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.Frame fraDept 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   340
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   2745
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   435
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   0
            Width           =   1905
         End
         Begin VB.Label lblDept 
            BackColor       =   &H80000005&
            Caption         =   "����"
            Height          =   285
            Left            =   0
            TabIndex        =   12
            Top             =   45
            Width           =   435
         End
      End
      Begin VB.Frame fraList 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   2880
         Left            =   60
         TabIndex        =   4
         Top             =   1125
         Width           =   2715
         Begin MSComctlLib.ListView lvwPati 
            Height          =   930
            Left            =   45
            TabIndex        =   7
            Top             =   225
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1640
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "סԺ��"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "�Ա�"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "��Ժʱ��"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "�ѱ�"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lvwDoc 
            Height          =   630
            Left            =   1425
            TabIndex        =   9
            Top             =   135
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   1111
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "�Ա�"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CommandButton cmdNone 
            Caption         =   "ȫ��"
            Height          =   330
            Left            =   855
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   2535
            Width           =   870
         End
         Begin VB.CommandButton cmdAll 
            Caption         =   "ȫѡ"
            Height          =   330
            Left            =   1785
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   2535
            Width           =   870
         End
         Begin MSComctlLib.ListView lvwDept 
            Height          =   870
            Left            =   1380
            TabIndex        =   8
            Top             =   960
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   1535
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lvwPatiOut 
            Height          =   930
            Left            =   90
            TabIndex        =   38
            Top             =   1260
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1640
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "סԺ��"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "�Ա�"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "��Ժʱ��"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "��Ժʱ��"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "�ѱ�"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fraDoc 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   340
         Left            =   15
         TabIndex        =   1
         Top             =   360
         Width           =   2715
         Begin VB.ComboBox cboDoc 
            Height          =   300
            Left            =   420
            TabIndex        =   2
            Text            =   "cboDoc"
            Top             =   0
            Width           =   1905
         End
         Begin VB.Label lblDoc 
            BackColor       =   &H80000005&
            Caption         =   "ҽ��"
            Height          =   225
            Left            =   0
            TabIndex        =   3
            Top             =   45
            Width           =   390
         End
      End
   End
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   4530
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   3390
      _Version        =   589884
      _ExtentX        =   5980
      _ExtentY        =   7990
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmMedRatioCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DoQuery(ByVal bytRan As Byte, ByVal bytWay As Byte, ByVal lngIDs As String, ByVal datBegin As Date, ByVal datEnd As Date)
Public Event CountWay(ByVal strWay As String, ByVal blnDrug As Boolean)

Private mstrPrivs As String

Private mstrDepParIDs As String
Private mstrDocParIDs As String
Private mstrPatParIDs As String

Private mstrPreDocDepID As String

Private mstrPrePatDepID As String
Private mstrPrePatDocID As String

Private mdatOutBegin As Date, mdatOutEnd As Date
Private mintOutPreTime As Integer '��һ��ѡ���ʱ���б��ֵ
Private mdatCurr As Date
Private mblnFirst As Boolean '��һ�γ���Ժ�����б�

Private Enum SeaRan '��ѯ��Χ
    ranȫԺ = 0
    ranסԺ = 2
    ran���� = 1
End Enum

Private Enum SeaWay '��ѯ��ʽ
    way�������� = 0
    way������ = 1
    way���� = 2
End Enum

Private Enum mCtlID
    opt_��������_��Ժ = 0
    opt_��������_��Ժ = 1
End Enum

Private Sub Form_Load()
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    Dim objPane As Pane
    Dim strDiffDate, strTmp As String
    Dim i, k As Integer
    Dim intIndex As Integer
    
    mstrPrivs = gstrPrivs
    Me.Width = tkpMain.Width: Me.Height = tkpMain.Height

    '����ؼ�------------------------------------------
    Call tkpMain.SetMargins(8, 8, 8, 8, 8)

    Set objGroup = tkpMain.Groups.Add(1, "�����б�")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picCond
    picCond.BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(2, "���������б�")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picDetail
    picDetail.BackColor = objItem.BackColor
    
    On Error GoTo errH
    i = Val(zlDatabase.GetPara("ҩ��ͳ�Ʒ�Χ", glngSys, 1261, "0"))
    optRan(i).Value = True
    
    If i = ran���� Then  '��ѯ��Χ�����ﲻ���ڰ����˲�ѯ
        optWay(way����).Enabled = False
        optWay(way����).Value = False
    Else
        optWay(way����).Enabled = True
    End If
    
    chkDrug.Value = Val(zlDatabase.GetPara("ҩƷ�ֱ�ͳ��", glngSys, 1261, 1))
 
    mdatCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    cboOutTime.Clear '��Ժ����ʱ�䷶Χ
    With cboOutTime
        .AddItem "������"
        .ItemData(.NewIndex) = 0
        .AddItem "������"
        .ItemData(.NewIndex) = 1
        .AddItem "ǰ����"
        .ItemData(.NewIndex) = 2
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "60����"
        .ItemData(.NewIndex) = 60
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboOutTime.ListCount > 0 Then cboOutTime.ListIndex = 0
    
    With cboTim
        .AddItem "����"
        .ItemData(.NewIndex) = 1
        .AddItem "���������"
        .ItemData(.NewIndex) = 2
        .AddItem "���һ����"
        .ItemData(.NewIndex) = 3
        .AddItem "���������"
        .ItemData(.NewIndex) = 6
        .AddItem "ָ��[...]"
        .ItemData(.NewIndex) = -1
    End With
    strDiffDate = zlDatabase.GetPara("ҩ�Ȳ�ѯ���", glngSys, 1261, "0")
    If IsNumeric(strDiffDate) Then
        Select Case strDiffDate
            Case "1", "0"
                cboTim.ListIndex = 0
            Case "2"
                cboTim.ListIndex = 1
            Case "3"
                cboTim.ListIndex = 2
            Case "6"
                cboTim.ListIndex = 3
        End Select
    Else
        Call Cbo.SetIndex(cboTim.hwnd, 4)
        dtpDate(0).Value = Format(Split(strDiffDate, "<Tab>")(0), "yyyy-MM-dd HH:mm")
        dtpDate(1).Value = Format(Split(strDiffDate, "<Tab>")(1), "yyyy-MM-dd HH:mm")
        dtpDate(0).Enabled = True
        dtpDate(1).Enabled = True
    End If
    
    mstrDepParIDs = zlDatabase.GetPara("ҩ�ȿ�������", glngSys, 1261, "")
    
    mstrDocParIDs = zlDatabase.GetPara("ҩ�ȿ�����", glngSys, 1261, "")
    If InStr(mstrDocParIDs, "|") > 0 Then
        mstrPreDocDepID = Split(mstrDocParIDs, "|")(0)
        mstrDocParIDs = Split(mstrDocParIDs, "|")(1)
    End If
    
    mstrPatParIDs = zlDatabase.GetPara("ҩ�Ȳ���", glngSys, 1261, "")
    If InStr(mstrPatParIDs, "|") > 0 Then
        strTmp = Split(mstrPatParIDs, "|")(0)
        If InStr(strTmp, ",") > 0 Then
            mstrPrePatDepID = Split(strTmp, ",")(0)
            mstrPrePatDocID = Split(strTmp, ",")(1)
        End If
        mstrPatParIDs = Split(mstrPatParIDs, "|")(1)
    End If
    
    i = Val(zlDatabase.GetPara("ҩ��ͳ�Ʒ�ʽ", glngSys, 1261, "0"))
    optWay(i).Value = True
    
    intIndex = i
    Call optWay_Click(intIndex)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboOutTime_Click()
'����ʱ�䷶Χ
    Dim intDateCount As Integer
    intDateCount = cboOutTime.ItemData(cboOutTime.ListIndex)
    
    If cboOutTime.ListIndex = mintOutPreTime And mintOutPreTime <> 6 Then Exit Sub
    If intDateCount = -1 Then
        If mdatOutBegin = CDate(0) Then
            mdatOutBegin = mdatCurr
            mdatOutEnd = mdatCurr
        End If
        If Not frmSelectTime.ShowMe(Me, mdatOutBegin, mdatOutEnd, cboOutTime) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
            Call Cbo.SetIndex(cboOutTime.hwnd, mintOutPreTime)
            Exit Sub
        End If
    Else
        mdatOutEnd = mdatCurr
        mdatOutBegin = mdatOutEnd - intDateCount
    End If
    
    If mdatOutBegin = CDate(0) Or mdatOutEnd = CDate(0) Then
        cboOutTime.ToolTipText = ""
    Else
        cboOutTime.ToolTipText = "��Χ��" & Format(mdatOutBegin, "yyyy-MM-dd") & " �� " & Format(mdatOutEnd, "yyyy-MM-dd")
    End If

    mintOutPreTime = cboOutTime.ListIndex
    
    Call LoadOutPati
End Sub

Private Sub LoadOutPati(Optional ByVal blnFirst As Boolean)
'���ܣ����س�Ժ����
'������blnFirst �Ƿ��ǵ�һ�γ��ֳ�Ժ�����б������δ�ҵ�����ʱ����ʾ
    Dim objListItem As ListItem
    Dim strSQL, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim strDocName, strPatParIDs, strPar As String
    Dim k As Integer
         
    strSQL = "Select a.����id,a.��ҳid,a.סԺ��,a.����,a.�Ա�,a.����,a.��Ժ����,a.��Ժ����,a.�ѱ� From ������ҳ A where a.��Ժ����id=[1]"
    
    If cboDoc.Text <> "����ҽ��" Then
        strSQL = strSQL & " And a.סԺҽʦ = [2]"
        strPar = Split(cboDoc.Text, "-")(1)
    End If
    
    strSQL = strSQL & " and a.��Ժ���� between [3] and [4]"
    
    strSQL = strSQL & "  Order By a.��Ժ���� desc"
 
    
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), strPar, mdatOutBegin, CDate(Format(mdatOutEnd, "YYYY-MM-DD 23:59:59")))
    
    lvwPatiOut.ListItems.Clear
    
    If rsTmp.EOF Then
        If Not blnFirst Then MsgBox "��ǰ������δ�ҵ���Ժ���ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    k = 0
    Do While Not rsTmp.EOF
        Set objListItem = lvwPatiOut.ListItems.Add(, "_" & rsTmp!����ID & "_" & rsTmp!��ҳID, "" & rsTmp!����)
            objListItem.SubItems(1) = "" & rsTmp!סԺ��
            objListItem.SubItems(2) = "" & rsTmp!�Ա�
            objListItem.SubItems(3) = "" & rsTmp!����
            objListItem.SubItems(4) = "" & rsTmp!��Ժ����
            objListItem.SubItems(5) = "" & rsTmp!��Ժ����
            objListItem.SubItems(6) = "" & rsTmp!�ѱ�
        rsTmp.MoveNext
    Loop
    Screen.MousePointer = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwDept Then
            Call cmdAll_Click
        ElseIf Me.ActiveControl Is lvwDoc Then
            Call cmdAll_Click
        ElseIf Me.ActiveControl Is lvwPati Then
            Call cmdAll_Click
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwDoc Then
            Call cmdNone_Click
        ElseIf Me.ActiveControl Is lvwPati Then
            Call cmdNone_Click
        ElseIf Me.ActiveControl Is lvwDept Then
            Call cmdNone_Click
        End If
    ElseIf KeyCode = 13 Then
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub LoadData(ByVal bytRan As Byte, ByVal bytWay As Byte)
'���ܣ��������ݣ��б�������б�
'������bytRan ��Χ,0-ȫԺ��1-���2-סԺ
'      bytPro �������� 0-���ʣ�1-ʵ��
'      bytWay ͳ�Ʒ�ʽ 0-���ң�1-�����ˣ�2-����
    Dim strSQL, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim objListItem As ListItem
    Dim strDepIDs As String
    Dim strDepParIDs As String
    Dim strDocName As String
    Dim strPreDeptID As String
    Dim i, k As Integer
    
    Screen.MousePointer = 11
    strSQL = "Select Distinct a.Id, a.����, a.����, a.���� From ���ű� A, ��������˵�� B Where a.Id = b.����id And b.��������='�ٴ�' And (a.����ʱ�� is NULL or Trunc(a.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
    If bytRan = 0 Then
        strTmp = " And b.������� <> 0"
    ElseIf bytRan = 1 Then
        strTmp = "  And b.������� = 1"
    ElseIf bytRan = 2 Then
        strTmp = " And (b.������� = 3 Or b.������� = 2)"
    End If
    If InStr(";" & mstrPrivs & ";", ";ȫԺ����;") = 0 Then
        strTmp = strTmp & " And a.Id in (Select t.����id From ������Ա T Where t.��Աid = [1])"
    End If
    
    strSQL = strSQL & strTmp & " Order By a.����"
    
    strDepParIDs = mstrDepParIDs
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    strDepIDs = GetUser����IDs(False)
    lvwDept.ListItems.Clear
    cboDept.Clear
    Do While Not rsTmp.EOF
        Set objListItem = lvwDept.ListItems.Add(, "_" & rsTmp!ID, "" & rsTmp!����)
            objListItem.SubItems(1) = "" & rsTmp!����
            objListItem.SubItems(2) = "" & rsTmp!����
            If InStr("," & strDepIDs & ",", "," & rsTmp!ID & ",") <> 0 And strDepParIDs = "" Then
                objListItem.Checked = True
                If k = 0 Then 'Ϊ�˿�����ѡ���
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
            If InStr("," & strDepParIDs & ",", "," & rsTmp!ID & ",") <> 0 Then
                objListItem.Checked = True
                If k = 0 Then 'Ϊ�˿�����ѡ���
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
            If bytWay <> 0 Then
                With cboDept
                    .AddItem rsTmp!���� & "-" & rsTmp!����
                    .ItemData(.NewIndex) = rsTmp!ID
                    If InStr("," & strDepIDs & ",", "," & rsTmp!ID & ",") <> 0 Then Call Cbo.SetIndex(cboDept.hwnd, .NewIndex)
                End With
            End If
        rsTmp.MoveNext
    Loop
    If cboDept.ListCount > 0 Then
        If cboDept.ListIndex = -1 Then Call Cbo.SetIndex(cboDept.hwnd, 0)
        If bytWay = 1 Then Call Cbo.Locate(cboDept, mstrPreDocDepID, True)
        If bytWay = 2 Then Call Cbo.Locate(cboDept, mstrPrePatDepID, True)
        Call cboDept_Click
    End If
    Screen.MousePointer = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboDept_Click()
'����ҽ���б�
    Dim objListItem As ListItem
    Dim strSQL, strTmp As String
    Dim strDocName, strDocParIDs As String
    Dim rsTmp As ADODB.Recordset
    Dim bytPrivDoc As Byte 'ҽ��Ȩ��
    Dim i, k As Integer
    
    For i = 0 To 2
        If optWay(i).Value Then Exit For
    Next i
    
    strDocParIDs = mstrDocParIDs
    strSQL = "Select a.Id, a.���, a.����, a.�Ա� From ��Ա�� A, ������Ա B, ��Ա����˵�� C" & vbNewLine & _
        "Where a.Id = b.��Աid And b.��Աid = c.��Աid And c.��Ա���� = 'ҽ��' And b.����id = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
        "Order By a.����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex))
    lvwDoc.ListItems.Clear
    Do While Not rsTmp.EOF
        Set objListItem = lvwDoc.ListItems.Add(, "_" & rsTmp!ID, "" & rsTmp!����)
            objListItem.SubItems(1) = "" & rsTmp!���
            objListItem.SubItems(2) = "" & rsTmp!�Ա�
            If UserInfo.ID = rsTmp!ID And strDocParIDs = "" Then
                objListItem.Checked = True
                If k = 0 Then 'Ϊ�˿�����ѡ���
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
            If InStr("," & strDocParIDs & ",", "," & rsTmp!ID & ",") <> 0 Then
                objListItem.Checked = True
                If k = 0 Then 'Ϊ�˿�����ѡ���
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
        rsTmp.MoveNext
    Loop
    If optWay(2).Value Then
        strSQL = "Select a.Id, a.���, a.����, a.�Ա�, c.��Ա����" & vbNewLine & _
            "From ��Ա�� A, ������Ա B, ��Ա����˵�� C" & vbNewLine & _
            "Where a.Id = b.��Աid And a.Id(+) = c.��Աid And b.����id = [1] And c.��Ա���� = 'ҽ��' Order By a.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex))
        cboDoc.Clear
        If InStr(";" & mstrPrivs & ";", ";ȫԺ����;") <> 0 Then
            With cboDoc
                .AddItem "����ҽ��"
                .ItemData(.NewIndex) = 0
            End With
            If InStr(";" & mstrPrivs & ";", ";���Ʋ���;") = 0 Then
                bytPrivDoc = 1
            End If
        End If
        Do While Not rsTmp.EOF
            With cboDoc
                .AddItem rsTmp!��� & "-" & rsTmp!����
                .ItemData(.NewIndex) = rsTmp!ID
                If UserInfo.ID = rsTmp!ID Then Call Cbo.SetIndex(cboDoc.hwnd, .NewIndex)
            End With
            rsTmp.MoveNext
        Loop
        If bytPrivDoc = 1 Then
            cboDoc.Clear
            With cboDoc
                .AddItem UserInfo.��� & "_" & UserInfo.����
                .ItemData(.NewIndex) = UserInfo.ID
            End With
        End If
        If cboDoc.ListCount > 0 Then
            If Not Cbo.Locate(cboDoc, mstrPrePatDocID, True) Then cboDoc.ListIndex = 0
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboDoc_Click()
'���ز����б�
    Dim objListItem As ListItem
    Dim strSQL, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim strDocName, strPatParIDs, strPar As String
    Dim k As Integer
    
    If optPatiType(opt_��������_��Ժ).Value And optPatiType(opt_��������_��Ժ).Enabled Then
        Call LoadOutPati
        Exit Sub
    End If
    
    strSQL = "Select a.����id, b.��ҳid, LPAD(a.��ǰ����,10,' ') as ��ǰ����,a.סԺ��, NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����, a.��Ժʱ��, a.�ѱ�" & vbNewLine & _
        "From ������Ϣ A, ������ҳ B, ��Ժ���� C" & vbNewLine & _
        "Where a.����id = c.����id And a.����id = b.����id And a.��ҳid = b.��ҳid And c.����id = [1]"
    
    If cboDoc.Text <> "����ҽ��" Then
        strSQL = strSQL & " And b.סԺҽʦ = [2]"
        strPar = Split(cboDoc.Text, "-")(1)
    End If
    strSQL = strSQL & "  Order By ��ǰ����"
    strPatParIDs = mstrPatParIDs
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), strPar)
    lvwPati.ListItems.Clear

    k = 0
    Do While Not rsTmp.EOF
        Set objListItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID & "_" & rsTmp!��ҳID, "" & rsTmp!����)
            objListItem.SubItems(1) = "" & rsTmp!��ǰ����
            objListItem.SubItems(2) = "" & rsTmp!סԺ��
            objListItem.SubItems(3) = "" & rsTmp!�Ա�
            objListItem.SubItems(4) = "" & rsTmp!����
            objListItem.SubItems(5) = "" & rsTmp!��Ժʱ��
            objListItem.SubItems(6) = "" & rsTmp!�ѱ�
            If UserInfo.���� = strDocName And strPatParIDs = "" Then
                objListItem.Checked = True
                If k = 0 Then 'Ϊ�˿�����ѡ���
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
            If InStr("," & strPatParIDs & ",", "," & rsTmp!����ID & ",") <> 0 Then
                objListItem.Checked = True
                If k = 0 Then 'Ϊ�˿�����ѡ���
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optPatiType_Click(Index As Integer)
    Call picDetail_Resize
    
    If Not mblnFirst And optPatiType(opt_��������_��Ժ).Value Then
        mblnFirst = True
        mdatOutBegin = mdatCurr
        mdatOutEnd = mdatCurr
        cboOutTime.ToolTipText = "��Χ��" & Format(mdatOutBegin, "yyyy-MM-dd") & " �� " & Format(mdatOutEnd, "yyyy-MM-dd")
        Call LoadOutPati(True)
    End If
End Sub

Private Sub optRan_Click(Index As Integer)
'���ܣ�ȷ��ʲô��ѯ��ʽ ���������� �������� ������
    Dim i As Integer
    Dim intTmp As Integer
    
    intTmp = -1
    
    If Index = ran���� Then
        If optWay(way����).Value Then
            optWay(way��������).Value = True
            intTmp = way��������
        End If
        optWay(way����).Enabled = False
        optWay(way����).Value = False
    Else
        optWay(way����).Enabled = True
        intTmp = way����
    End If
    
    If intTmp = -1 Then intTmp = IIf(optWay(way��������).Value, way��������, way������)
    
    Call LoadData(Index, intTmp)
End Sub

Private Sub cmdSearch_Click()
    Dim i As Integer
    Dim intCount As Integer
    Dim strIDs As String
    Dim bytRan As Byte '���÷�Χ
    Dim bytPro As Byte '1-���ʽ�2-ʵ�ս��
    Dim blnDebt As Boolean '�������۵�
    Dim bytWay As Byte 'ͳ�Ʒ�ʽ
    Dim strNow As String
    Dim strWorkTim As String
    Dim strTmp As String
    Dim objLvw As Object

    strIDs = ""
    
    On Error GoTo errH
    
    strWorkTim = zlDatabase.GetPara("�������°�ʱ��", glngSys)
    
    If strWorkTim = "" Then strWorkTim = "08:00 AND 12:00"
    
    strNow = Format(zlDatabase.Currentdate, "hh:mm")

    If Split(strWorkTim, " AND ")(0) < strNow And Split(strWorkTim, " AND ")(1) > strNow And Not optWay(2).Value Then
        MsgBox "Ŀǰ���������ϰ�ʱ�䣬�������������Ҳ�ѯ�Ͱ������˲�ѯ��", vbInformation, gstrSysName
        
        '��ѯ��Χ���������������ù�������
        If Not optRan(ran����).Value Then optWay(way����).Value = True
        
        Exit Sub
    End If
    
    If optRan(ranȫԺ).Value Then
        bytRan = 0 'ȫԺ
    ElseIf optRan(ranסԺ).Value Then
        bytRan = 2 'סԺ
    ElseIf optRan(ran����).Value Then
        bytRan = 1 '����
    End If
    
    If lvwPati.Visible Or lvwPatiOut.Visible Then
        If lvwPati.Visible Then
            Set objLvw = lvwPati
        Else
            Set objLvw = lvwPatiOut
        End If
                
        With objLvw
            For i = 1 To .ListItems.Count
                If .ListItems(i).Checked Then
                    strIDs = strIDs & "," & Split(.ListItems(i).Key, "_")(1) & ":" & Split(.ListItems(i).Key, "_")(2)
                    strTmp = strTmp & "," & Split(.ListItems(i).Key, "_")(1)
                End If
            Next
        End With
        
        bytWay = 2
        If strIDs = "" Then MsgBox "��ѡ����", vbInformation, gstrSysName: Exit Sub
        mstrPrePatDepID = cboDept.ItemData(cboDept.ListIndex)
        mstrPrePatDocID = cboDoc.ItemData(cboDoc.ListIndex)
        mstrPatParIDs = Mid(strTmp, 2)
    ElseIf lvwDept.Visible Then
        For i = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(i).Checked Then
                strIDs = strIDs & "," & Split(lvwDept.ListItems(i).Key, "_")(1)
            End If
        Next
        bytWay = 0
        If strIDs = "" Then MsgBox "��ѡ�����", vbInformation, gstrSysName: Exit Sub
        mstrDepParIDs = Mid(strIDs, 2)
    ElseIf lvwDoc.Visible Then
        For i = 1 To lvwDoc.ListItems.Count
            If lvwDoc.ListItems(i).Checked Then
                strIDs = strIDs & "," & Split(lvwDoc.ListItems(i).Key, "_")(1)
            End If
        Next
        bytWay = 1
        If strIDs = "" Then MsgBox "��ѡ�񿪵���", vbInformation, gstrSysName: Exit Sub
        mstrPreDocDepID = cboDept.ItemData(cboDept.ListIndex)
        mstrDocParIDs = Mid(strIDs, 2)
    End If
    
    strIDs = Mid(strIDs, 2)
    
    Call zlDatabase.SetPara("ҩ��ͳ�Ʒ�Χ", bytRan, glngSys, 1261)
    Call zlDatabase.SetPara("ҩƷ�ֱ�ͳ��", chkDrug.Value, glngSys, 1261)
'    Call zlDatabase.SetPara("ҩ�Ⱥ����۵�", chkNotPay.Value, glngSys, 1261)
    Call zlDatabase.SetPara("ҩ��ͳ�Ʒ�ʽ", bytWay, glngSys, 1261)
    
    If cboTim.ItemData(cboTim.ListIndex) = -1 Then
        Call zlDatabase.SetPara("ҩ�Ȳ�ѯ���", dtpDate(0).Value & "<Tab>" & dtpDate(1).Value, glngSys, 1261)
    Else
        intCount = cboTim.ItemData(cboTim.ListIndex)
        Call zlDatabase.SetPara("ҩ�Ȳ�ѯ���", intCount, glngSys, 1261)
    End If
    
    RaiseEvent DoQuery(bytRan, bytWay, strIDs, dtpDate(0).Value, dtpDate(1).Value)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errH
    
    Call zlDatabase.SetPara("ҩ�ȿ�������", mstrDepParIDs, glngSys, 1261)
    mstrDocParIDs = mstrPreDocDepID & "|" & mstrDocParIDs
    Call zlDatabase.SetPara("ҩ�ȿ�����", mstrDocParIDs, glngSys, 1261)
    mstrPatParIDs = mstrPrePatDepID & "," & mstrPrePatDocID & "|" & mstrPatParIDs
    Call zlDatabase.SetPara("ҩ�Ȳ���", mstrPatParIDs, glngSys, 1261)
    
    mstrPrivs = ""
    mstrDepParIDs = ""
    mstrDocParIDs = ""
    mstrPatParIDs = ""
    mstrPreDocDepID = ""
    mstrPrePatDepID = ""
    mstrPrePatDocID = ""
    mblnFirst = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optWay_Click(Index As Integer)
'���ܣ���ѡͳ�Ʒ�ʽ  ���������ң��������ˣ�������
    Dim strWay As String
    Dim i As Integer
    
    Select Case Index
        Case 0
            strWay = "��������"
            tkpMain.Groups.Find(2).Caption = "���������б�"
        Case 1
            strWay = "������"
            tkpMain.Groups.Find(2).Caption = "�������б�"
        Case 2
            strWay = "����"
            tkpMain.Groups.Find(2).Caption = "�����б�"
    End Select
    
    For i = 0 To 2
        If optRan(i).Value Then Exit For
    Next i
    
    Call LoadData(i, Index)
    
    Call picDetail_Resize
    
    RaiseEvent CountWay(strWay, chkDrug.Value = 1)
    
End Sub

Private Sub picDetail_Resize()
    
    On Error Resume Next
    
    If optWay(way��������).Value Then
            
        lvwPati.Visible = False
        lvwDoc.Visible = False
        lvwDept.Visible = True
        
        fraDept.Visible = False
        fraDoc.Visible = False
        fraList.Top = 0
        picPatiType.Enabled = False
        optPatiType(opt_��������_��Ժ).Enabled = False
        optPatiType(opt_��������_��Ժ).Enabled = False
        
    ElseIf optWay(way������).Value Then
            
        lvwPati.Visible = False
        lvwDoc.Visible = True
        lvwDept.Visible = False
        
        fraDoc.Visible = False
        fraDept.Visible = True
        fraDept.Left = 0
        fraDept.Top = 0
        fraList.Top = 350
        picPatiType.Enabled = False
        optPatiType(opt_��������_��Ժ).Enabled = False
        optPatiType(opt_��������_��Ժ).Enabled = False
    ElseIf optWay(way����).Value Then
        
        lvwPati.Visible = True
        lvwDoc.Visible = False
        lvwDept.Visible = False
        
        fraDept.Visible = True
        fraDoc.Visible = True
        fraDept.Left = 0
        fraDoc.Left = 0
        fraDept.Top = 0
        fraDoc.Top = fraDept.Height
        fraList.Top = fraDept.Height + fraDoc.Height + 10
        
        picPatiType.Enabled = True
        optPatiType(opt_��������_��Ժ).Enabled = True
        optPatiType(opt_��������_��Ժ).Enabled = True
    End If
    
    If optWay(way����).Value Then
        cboTim.Enabled = False
        dtpDate(0).Enabled = False
        dtpDate(1).Enabled = False
    Else
        cboTim.Enabled = True
        With cboTim
            If .ItemData(.ListIndex) <> -1 Then
                dtpDate(0).Enabled = False
                dtpDate(1).Enabled = False
            Else
                dtpDate(0).Enabled = True
                dtpDate(1).Enabled = True
            End If
        End With
    End If
    
    If picPatiType.Enabled Then
        If optPatiType(opt_��������_��Ժ).Value Then
            fraOutTime.Visible = True
            lvwPati.Visible = False
            lvwPatiOut.Visible = True
            fraOutTime.Width = picDetail.Width
            fraOutTime.Left = 0
            cboOutTime.Width = 1530
            fraOutTime.Top = fraDoc.Top + fraDoc.Height
            fraOutTime.Height = 350
            fraList.Top = fraOutTime.Top + fraOutTime.Height
        Else
            fraOutTime.Visible = False
            lvwPatiOut.Visible = False
            fraList.Top = fraDoc.Top + fraDoc.Height
        End If
    Else
        fraOutTime.Visible = False
        lvwPatiOut.Visible = False
    End If
    
    fraList.Left = 0
    fraList.Width = Me.ScaleWidth
    fraList.Height = Me.ScaleHeight
 
    lvwPati.Left = 0
    lvwPati.Width = fraList.Width - 800
    lvwPati.Height = 2490
    lvwPati.Top = 0
    
    
    lvwPatiOut.Left = 0
    lvwPatiOut.Width = fraList.Width - 800
    lvwPatiOut.Height = 2490
    lvwPatiOut.Top = 0
    
 
    lvwDept.Left = 0
    lvwDept.Width = fraList.Width - 800
    lvwDept.Height = 2490
    lvwDept.Top = 0
 
    lvwDoc.Left = 0
    lvwDoc.Width = fraList.Width - 800
    lvwDoc.Height = 2490
    lvwDoc.Top = 0
 
    cmdAll.Left = lvwDoc.Width - cmdAll.Width - 100
    cmdNone.Left = cmdAll.Left - 30 - cmdAll.Width
    cmdNone.Top = cmdAll.Top
End Sub

Private Sub picCond_Resize()
    On Error Resume Next
    
    picCond.Height = picCond.Height
    fraRange.Width = picCond.ScaleWidth - fraRange.Left
    fraRange.Height = 580
    fraRange.Top = 0
    
    fraWay.Width = fraRange.Width
    chkDrug.Width = fraWay.Width
    dtpDate(0).Width = picCond.ScaleWidth - dtpDate(0).Left
    dtpDate(1).Width = dtpDate(0).Width
    dtpDate(2).Width = dtpDate(0).Width
    cboTim.Width = dtpDate(0).Width
    cmdSearch.Width = cmdAll.Width
    cmdSearch.Left = fraRange.Width + fraRange.Left - cmdAll.Width
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tkpMain.Left = 0
    tkpMain.Top = 0
    tkpMain.Width = Me.ScaleWidth
    tkpMain.Height = Me.ScaleHeight
End Sub

Private Sub chkDrug_Click()
    Dim strWay As String
    
    If optWay(0).Value Then
        strWay = "��������"
    ElseIf optWay(1).Value Then
        strWay = "������"
    ElseIf optWay(2).Value Then
        strWay = "����"
    End If
    
    RaiseEvent CountWay(strWay, chkDrug.Value = 1)
End Sub

Private Sub cboTim_Click()
    Dim curDate As Date
    
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    With cboTim
        If .ItemData(.ListIndex) <> -1 Then
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
            dtpDate(0).Value = Format(DateAdd("m", -1 * .ItemData(.ListIndex) + 1, curDate), "yyyy-MM-1 00:00")
            dtpDate(1).Value = Format(curDate, "yyyy-MM-dd HH:mm")
        Else
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
            dtpDate(0).SetFocus
            dtpDate(0).Value = Format(curDate, "yyyy-MM-1 00:00")
            dtpDate(1).Value = Format(curDate, "yyyy-MM-dd HH:mm")
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdNone_Click()
    If lvwPati.Visible Then
        Call SelectLVW(lvwPati, False)
    ElseIf lvwPatiOut.Visible Then
        Call SelectLVW(lvwPatiOut, False)
    ElseIf lvwDept.Visible Then
        Call SelectLVW(lvwDept, False)
    ElseIf lvwDoc.Visible Then
        Call SelectLVW(lvwDoc, False)
    End If
End Sub

Private Sub cmdAll_Click()
    If lvwPati.Visible Then
        Call SelectLVW(lvwPati, True)
    ElseIf lvwPatiOut.Visible Then
        Call SelectLVW(lvwPatiOut, True)
    ElseIf lvwDept.Visible Then
        Call SelectLVW(lvwDept, True)
    ElseIf lvwDoc.Visible Then
        Call SelectLVW(lvwDoc, True)
    End If
End Sub

Private Sub SelectLVW(objLvw As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    
    For i = 1 To objLvw.ListItems.Count
        objLvw.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub lvwDept_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwDept, ColumnHeader.Index)
End Sub

Private Sub lvwDoc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwDoc, ColumnHeader.Index)
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwDept_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = IIf(Item.Checked, False, True)
End Sub

Private Sub lvwDoc_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = IIf(Item.Checked, False, True)
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = IIf(Item.Checked, False, True)
End Sub
