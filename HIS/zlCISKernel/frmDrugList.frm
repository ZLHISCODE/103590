VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugList 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "������ҩ�嵥"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18540
   Icon            =   "frmDrugList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10185
   ScaleWidth      =   18540
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picAdviceFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   2040
      ScaleHeight     =   2625
      ScaleWidth      =   2985
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmdAdviceQuit 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   300
         Left            =   1560
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdviceOK 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "ȷ��(&O)"
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkZY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��ҩ"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   1680
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkXY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��ҩ"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1680
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkסԺ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "סԺ"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   1230
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkMZ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1230
         Value           =   1  'Checked
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   300
         Left            =   975
         TabIndex        =   6
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   201981955
         CurrentDate     =   40976
      End
      Begin MSComCtl2.DTPicker dtpStopTime 
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   201981955
         CurrentDate     =   40976
      End
      Begin VB.Label lblFilter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��ҩ�Զ���ȡ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   960
         TabIndex        =   22
         Top             =   60
         Width           =   1560
      End
      Begin VB.Image imgIco 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   600
         Picture         =   "frmDrugList.frx":6852
         Stretch         =   -1  'True
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblAdvice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ���ࣺ"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1777
         Width           =   900
      End
      Begin VB.Label lblAdvice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����Դ��"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1327
         Width           =   900
      End
      Begin VB.Label lblAdvice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ�䣺"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   900
         Width           =   900
      End
      Begin VB.Label lblAdvice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ�䣺"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   420
         Width           =   900
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C000&
      Height          =   465
      Left            =   720
      ScaleHeight     =   465
      ScaleWidth      =   10455
      TabIndex        =   14
      Top             =   1200
      Width           =   10455
      Begin VB.PictureBox picTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6480
         ScaleHeight     =   375
         ScaleWidth      =   3975
         TabIndex        =   23
         Top             =   38
         Width           =   3975
         Begin VB.OptionButton OptTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   2
            Top             =   0
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton OptTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "��һ��"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   3
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton OptTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "������"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   2160
            TabIndex        =   4
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton OptTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "������"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   3120
            TabIndex        =   5
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblTime 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   24
            Top             =   90
            Width           =   360
         End
      End
      Begin VB.Label lblPati 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1050
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   9825
      Width           =   18540
      _ExtentX        =   32703
      _ExtentY        =   635
      SimpleText      =   $"frmDrugList.frx":D0A4
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugList.frx":D0EB
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   27623
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   5925
      _cx             =   10451
      _cy             =   6271
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
      MouseIcon       =   "frmDrugList.frx":D97F
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   400
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   10000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugList.frx":E259
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox pictmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1920
         ScaleHeight     =   240
         ScaleWidth      =   480
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   360
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDrugList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnMod As Boolean   'true ����ģ̬��ʾ��false��ģ̬��ʾ
Private mlng����ID As Long   '����id
Private mlng��ҳID As Long   '��ҳid
Private mlngEditTag As Long  '����״̬���ã�0-����״̬��1-�༭״̬
Private mblnReturn As Boolean
Private mstrLike As String
Private mint���� As Integer
Private mstrTip As String
Private mlngLastColor As Long   '�ϴ�ѡ������ɫ

Private Type PointAPI
        X As Long
        Y As Long
End Type

Private Const GRD_UNEDITCELL_COLOR = &H8000000B  'δ�༭�ĵ�Ԫ����ɫ������ɫ
Private Const Red_COLOR = &HC0C0FF  '����ɫ


Private Enum COL��ҩ�嵥
    '������
    COL_ID = 1
    COL_����ID = 2
    COL_��ҳID = 3
    col_��� = 4
    COL_��ҩ��Դ = 5
    COL_������ĿID = 6
    COL_�շ�ϸĿID = 7
    COL_Ƶ�ʼ�� = 8
    COL_�����λ = 9
    COL_�÷�id = 10
    col_�巨id = 11
    COL_��ֹʱ�� = 12
    '�ɼ���
    COL_��ʼʱ�� = 13
    col_ҩƷ��� = 14
    col_��ҩ���� = 15
    COL_�÷� = 16
    COL_�������� = 17
    COL_������λ = 18
    COL_�ܸ����� = 19
    COL_������λ = 20
    COL_���� = 21
    COL_ִ��Ƶ�� = 22
    COL_��ע = 23
    
    '������
    COL_�Ǽ��� = 24
    COL_�Ǽ�ʱ�� = 25
    col_�䷽���� = 26 '��ʽ:[�䷽����]��ҩ����<Data>������ĿID<Data>�շ�ϸĿID<Data>����<Data>��ע<Data>��λ
    col_�Ƿ��޸� = 27
    COL_Ƶ�ʴ��� = 28
End Enum

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Private Function LoadDrug()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intTime As Integer
    Dim i As Long
    
    On Error GoTo errH
    With vsAdvice
        For i = 0 To OptTime.Count - 1
            If OptTime(i).value = True Then
                intTime = decode(OptTime(i).Caption, "��һ��", 1, "������", 3, "������", 6)
                Exit For
            End If
        Next
        strSQL = "Select a.Id, a.����id, a.��ҳid, a.���, a.��ҩ��Դ, a.ҩƷ���, a.��ҩ����, a.������Ŀid, a.�շ�ϸĿid, a.����, a.��ʼʱ��, a.��ֹʱ��, a.�Ǽ�ʱ��, a. �Ǽ���,a. �ܸ�����, a.��������, a. ִ��Ƶ��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, a.�÷�id, a.�巨id, a.��ע, b.���㵥λ,C.���� as �÷�,D.סԺ��λ" & _
                " From ������ҩ�嵥 A,������ĿĿ¼ B, ������ĿĿ¼ C, ҩƷ��� D Where a.������Ŀid = b.Id(+) And a.�շ�ϸĿid=D.ҩƷID(+) And A.�÷�ID=C.ID(+) And a.����id = [1]" & IIF(intTime = 0, "", " And a.��ʼʱ�� Between add_months(sysdate,-[2]) And Sysdate") & " Order By a.��ʼʱ��,a.���,a.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, intTime)
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '������
                .TextMatrix(i, COL_ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COL_����ID) = Val(rsTmp!����ID & "")
                .TextMatrix(i, COL_��ҳID) = Val(rsTmp!��ҳID & "")
                .TextMatrix(i, col_���) = Val(rsTmp!��� & "")
                .TextMatrix(i, COL_��ҩ��Դ) = Val(rsTmp!��ҩ��Դ & "")
                .TextMatrix(i, COL_������ĿID) = Val(rsTmp!������ĿID & "")
                .TextMatrix(i, COL_�շ�ϸĿID) = Val(rsTmp!�շ�ϸĿID & "")
                .TextMatrix(i, COL_Ƶ�ʼ��) = Val(rsTmp!Ƶ�ʼ�� & "")
                .TextMatrix(i, COL_�����λ) = rsTmp!�����λ & ""
                .TextMatrix(i, COL_�÷�id) = Val(rsTmp!�÷�ID & "")
                .TextMatrix(i, col_�巨id) = Val(rsTmp!�巨ID & "")
                .TextMatrix(i, COL_��ֹʱ��) = Format(rsTmp!��ֹʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_��ʼʱ��) = Format(rsTmp!��ʼʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, col_ҩƷ���) = decode(rsTmp!ҩƷ��� & "", "5", "����ҩ", "6", "�г�ҩ", "�в�ҩ")
                .TextMatrix(i, col_��ҩ����) = rsTmp!��ҩ���� & ""
                .TextMatrix(i, COL_�÷�) = rsTmp!�÷� & ""
                .TextMatrix(i, COL_��������) = IIF(.TextMatrix(i, col_ҩƷ���) = "�в�ҩ", "", FormatEx(NVL(rsTmp!��������), 5))
                .TextMatrix(i, COL_�ܸ�����) = FormatEx(NVL(rsTmp!�ܸ�����), 5)
                .TextMatrix(i, COL_������λ) = IIF(.TextMatrix(i, col_ҩƷ���) = "�в�ҩ", "", rsTmp!���㵥λ & "")
                .TextMatrix(i, COL_������λ) = IIF(.TextMatrix(i, col_ҩƷ���) = "�в�ҩ", "��", rsTmp!סԺ��λ & "")
                .TextMatrix(i, COL_����) = FormatEx(NVL(rsTmp!����), 5)
                .TextMatrix(i, COL_ִ��Ƶ��) = rsTmp!ִ��Ƶ�� & ""
                .TextMatrix(i, COL_��ע) = rsTmp!��ע & ""
                .TextMatrix(i, COL_�Ǽ���) = rsTmp!�Ǽ��� & ""
                .TextMatrix(i, COL_�Ǽ�ʱ��) = Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-mm-dd hh:mm")
                
                '��������
                .Cell(flexcpData, i, COL_ִ��Ƶ��) = .TextMatrix(i, COL_ִ��Ƶ��)
                .Cell(flexcpData, i, COL_�÷�) = .TextMatrix(i, COL_�÷�)
                .Cell(flexcpData, i, col_��ҩ����) = .TextMatrix(i, col_��ҩ����)
                .Cell(flexcpData, i, col_ҩƷ���) = decode(.TextMatrix(i, col_ҩƷ���), "����ҩ", "5", "�г�ҩ", "6", "�в�ҩ", "8")
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
        Else
            .Rows = .FixedRows + 1
        End If
        .WordWrap = True
        '�Զ������и�
        .AutoSize col_��ҩ����
        .Cell(flexcpBackColor, .FixedRows, col_ҩƷ���, .Rows - 1, col_ҩƷ���) = GRD_UNEDITCELL_COLOR      '����ɫ
        .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
        .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, 0) = GRD_UNEDITCELL_COLOR
        SetTagһ����ҩ
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function SaveDrug()
    Dim lngID As Long
    Dim i As Long, j As Long
    Dim arrSQL As Variant
    Dim dtNow As Date
    Dim blnTran As Boolean
    Dim arrTime As Variant, arrTmp As Variant
    Dim lng��� As Long
    
    arrSQL = Array()
    On Error GoTo errH
    dtNow = zlDatabase.Currentdate
    With vsAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col_��ҩ����) <> "" Then
                lng��� = 0
                If Val(.TextMatrix(i, COL_ID)) <> 0 Then
                    If .TextMatrix(i, col_�Ƿ��޸�) = "1" Then
                        If .RowHidden(i) = True Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_������ҩ�嵥_Delete(" & Val(.TextMatrix(i, COL_ID)) & ")"
                        Else
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            
                            arrSQL(UBound(arrSQL)) = "Zl_������ҩ�嵥_Update(" & Val(.TextMatrix(i, COL_ID)) & "," & ZVal(.TextMatrix(i, col_���)) & "," & Val(.TextMatrix(i, COL_��ҩ��Դ)) & ",'" & _
                                Val(.Cell(flexcpData, i, col_ҩƷ���)) & "','" & .TextMatrix(i, col_��ҩ����) & "'," & ZVal(.TextMatrix(i, COL_������ĿID)) & "," & ZVal(.TextMatrix(i, COL_�շ�ϸĿID)) & "," & _
                                ZVal(.TextMatrix(i, COL_����)) & ",To_Date('" & Format(.TextMatrix(i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & IIF(.TextMatrix(i, COL_��ֹʱ��) = "", "Null", "To_Date('" & Format(.TextMatrix(i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')") & "," & _
                                "To_Date('" & Format(dtNow, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & .TextMatrix(i, COL_�Ǽ���) & "'," & ZVal(.TextMatrix(i, COL_�ܸ�����)) & "," & ZVal(.TextMatrix(i, COL_��������)) & ",'" & _
                                .TextMatrix(i, COL_ִ��Ƶ��) & "'," & ZVal(.TextMatrix(i, COL_Ƶ�ʴ���)) & "," & ZVal(.TextMatrix(i, COL_Ƶ�ʼ��)) & ",'" & .TextMatrix(i, COL_�����λ) & "'," & ZVal(.TextMatrix(i, COL_�÷�id)) & "," & _
                                ZVal(.TextMatrix(i, col_�巨id)) & ",'" & .TextMatrix(i, COL_��ע) & "')"
                        
                            If .TextMatrix(i, col_ҩƷ���) = "�в�ҩ" Or .Cell(flexcpData, i, col_�䷽����) <> "" Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_������ҩ�䷽_Delete(" & Val(.TextMatrix(i, COL_ID)) & ")"
                            End If
                            If .TextMatrix(i, col_ҩƷ���) = "�в�ҩ" Then
                                arrTime = Split(.TextMatrix(i, col_�䷽����), "[�䷽����]")
                                For j = 1 To UBound(arrTime)
                                    arrTmp = Split(arrTime(j), "<Data>")
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = "Zl_������ҩ�䷽_Insert(" & Val(.TextMatrix(i, COL_ID)) & "," & j & "," & Val(arrTmp(1)) & "," & Val(arrTmp(2)) & "," & ZVal(arrTmp(3)) & ",'" & arrTmp(4) & "')"
                                Next
                            End If
                        End If
                    End If
                Else
                    lngID = zlDatabase.GetNextID("������ҩ�嵥")
                    
                    'ת��һ����ҩID
                    If .TextMatrix(i, col_ҩƷ���) <> "�в�ҩ" And Val(.TextMatrix(i, col_���)) <> 0 Then
                        If Val(.TextMatrix(i, col_���)) < 0 Then
                            If Val(.TextMatrix(i, col_���)) = Val(.TextMatrix(Abs(Val(.TextMatrix(i, col_���))), col_���)) Then
                                If i = Abs(Val(.TextMatrix(i, col_���))) Then
                                   lng��� = lngID
                                Else
                                   lng��� = Val(.Cell(flexcpData, Abs(Val(.TextMatrix(i, col_���))), col_���))
                                End If
                            End If
                            .Cell(flexcpData, i, col_���) = lng���
                        Else
                            lng��� = Val(.TextMatrix(i, col_���))
                        End If
                    End If
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_������ҩ�嵥_Insert(" & lngID & "," & mlng����ID & "," & mlng��ҳID & "," & ZVal(lng���) & "," & Val(.TextMatrix(i, COL_��ҩ��Դ)) & ",'" & _
                        Val(.Cell(flexcpData, i, col_ҩƷ���)) & "','" & .TextMatrix(i, col_��ҩ����) & "'," & ZVal(.TextMatrix(i, COL_������ĿID)) & "," & ZVal(.TextMatrix(i, COL_�շ�ϸĿID)) & "," & _
                        ZVal(.TextMatrix(i, COL_����)) & ",To_Date('" & Format(.TextMatrix(i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & IIF(.TextMatrix(i, COL_��ֹʱ��) = "", "Null", "To_Date('" & Format(.TextMatrix(i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')") & "," & _
                        "To_Date('" & Format(dtNow, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & .TextMatrix(i, COL_�Ǽ���) & "'," & ZVal(.TextMatrix(i, COL_�ܸ�����)) & "," & ZVal(.TextMatrix(i, COL_��������)) & ",'" & _
                        .TextMatrix(i, COL_ִ��Ƶ��) & "'," & ZVal(.TextMatrix(i, COL_Ƶ�ʴ���)) & "," & ZVal(.TextMatrix(i, COL_Ƶ�ʼ��)) & ",'" & .TextMatrix(i, COL_�����λ) & "'," & ZVal(.TextMatrix(i, COL_�÷�id)) & "," & _
                        ZVal(.TextMatrix(i, col_�巨id)) & ",'" & .TextMatrix(i, COL_��ע) & "')"
                        
                    If .TextMatrix(i, col_ҩƷ���) = "�в�ҩ" Then
                        arrTime = Split(.TextMatrix(i, col_�䷽����), "[�䷽����]")
                        For j = 1 To UBound(arrTime)
                            arrTmp = Split(arrTime(j), "<Data>")
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_������ҩ�䷽_Insert(" & lngID & "," & j & "," & Val(arrTmp(1)) & "," & Val(arrTmp(2)) & "," & ZVal(arrTmp(3)) & ",'" & arrTmp(4) & "')"
                        Next
                    End If
                    .Cell(flexcpData, i, COL_�շ�ϸĿID) = lngID  '����id
                End If
                
            End If
        Next
    End With
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    Screen.MousePointer = 0
    
    SaveDrug = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function





Private Sub UpdateDrug()
    '���ܣ����µ�ǰ����
    Dim i As Long, lngRows As Long

    With vsAdvice
        lngRows = .Rows - 1
        For i = lngRows To 1 Step -1
            If .TextMatrix(i, col_�Ƿ��޸�) = "1" Then .TextMatrix(i, col_�Ƿ��޸�) = ""
            If Val(.Cell(flexcpData, i, col_���)) > 0 And Val(.TextMatrix(i, COL_ID)) = 0 And Val(.TextMatrix(i, col_���)) < 0 Then .TextMatrix(i, col_���) = Val(.Cell(flexcpData, i, col_���))
            .Cell(flexcpData, i, col_���) = ""
            If .Cell(flexcpData, i, COL_�շ�ϸĿID) <> "" Then .TextMatrix(i, COL_ID) = Val(.Cell(flexcpData, i, COL_�շ�ϸĿID)): .Cell(flexcpData, i, COL_�շ�ϸĿID) = ""
            .Cell(flexcpBackColor, i, 0, i, 0) = GRD_UNEDITCELL_COLOR
            If .RowHidden(i) = True Then .RemoveItem (i)
        Next
        vsAdvice.Tag = ""
    End With
End Sub


Public Function ShowMe(frmParent As Object, ByVal blnMod As Boolean, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ�������ҩ�嵥
'������frmParent ������
'      blnMod �Ƿ���ģ̬��ʽ��ʾ
'      lng����ID,
'      lng��ҳID,
'���أ�
    mblnMod = blnMod
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    Me.Show IIF(blnMod, 1, 0), frmParent
End Function

Private Function checkDrug()
    Dim i As Long, j As Long
    
    mstrTip = ""
    With vsAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col_��ҩ����) <> "" And .RowHidden(i) = False Then
                '�ָ���ɫ
                If .Cell(flexcpBackColor, i, COL_�÷�, i, COL_�÷�) = Red_COLOR Then .Cell(flexcpBackColor, i, COL_�÷�, i, COL_�÷�) = 0
                If .Cell(flexcpBackColor, i, col_��ҩ����, i, col_��ҩ����) = Red_COLOR Then .Cell(flexcpBackColor, i, col_��ҩ����, i, col_��ҩ����) = 0
                If .Cell(flexcpBackColor, i, COL_��ʼʱ��, i, COL_��ʼʱ��) = Red_COLOR Then .Cell(flexcpBackColor, i, COL_��ʼʱ��, i, COL_��ʼʱ��) = 0
                
                If .TextMatrix(i, COL_��ʼʱ��) = "" Then
                    .Cell(flexcpBackColor, i, COL_��ʼʱ��, i, COL_��ʼʱ��) = Red_COLOR
                    MsgBox "��ҩ�嵥�Ŀ�ʼʱ��Ϊ������,��¼�롣", vbInformation, gstrSysName
                    mstrTip = i & "|" & COL_��ʼʱ�� & "|" & "��ҩ�嵥��ʼʱ��Ϊ������,��¼�롣"
                    .Row = i: .Col = COL_��ʼʱ��: Call vsAdvice.ShowCell(.Row, .Col)
                    Exit Function
                End If
                
                If .TextMatrix(i, COL_�÷�) = "" Then
                    .Cell(flexcpBackColor, i, COL_�÷�, i, COL_�÷�) = Red_COLOR
                    MsgBox "��ҩ�嵥���÷�Ϊ������,��¼�롣", vbInformation, gstrSysName
                    mstrTip = i & "|" & COL_�÷� & "|" & "��ҩ�嵥���÷�Ϊ������,��¼�롣"
                    .Row = i: .Col = COL_�÷�: Call vsAdvice.ShowCell(.Row, .Col)
                    Exit Function
                End If
                
                If i <> .Rows - 1 And .TextMatrix(i, col_ҩƷ���) <> "�в�ҩ" Then '����Ƿ������ͬ��ҩ�嵥
                    For j = .Rows - 1 To i + 1 Step -1
                        If .Cell(flexcpBackColor, i, col_��ҩ����, i, col_��ҩ����) = Red_COLOR Then .Cell(flexcpBackColor, i, col_��ҩ����, i, col_��ҩ����) = 0
                        If .TextMatrix(j, col_��ҩ����) <> "" And .RowHidden(j) = False Then
                            If .TextMatrix(j, COL_��ʼʱ��) & "|" & .TextMatrix(j, col_��ҩ����) & "|" & .TextMatrix(j, col_ҩƷ���) & "|" & .TextMatrix(j, COL_������ĿID) & "|" & .TextMatrix(j, COL_�շ�ϸĿID) = .TextMatrix(i, COL_��ʼʱ��) & "|" & .TextMatrix(i, col_��ҩ����) & "|" & .TextMatrix(i, col_ҩƷ���) & "|" & .TextMatrix(i, COL_������ĿID) & "|" & .TextMatrix(i, COL_�շ�ϸĿID) Then
                                .Cell(flexcpBackColor, j, col_��ҩ����, j, col_��ҩ����) = Red_COLOR
                                MsgBox "���������ظ�����ҩ�嵥,���顣", vbInformation, gstrSysName
                                mstrTip = j & "|" & col_��ҩ���� & "|" & "���������ظ�����ҩ�嵥,���顣"
                                .Row = j: .Col = col_��ҩ����: Call vsAdvice.ShowCell(.Row, .Col)
                                Exit Function
                            End If
                        End If
                    Next
                End If
            End If
        Next
        checkDrug = True
    End With
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim Pt As PointAPI
    Dim i As Long, lngTmp As Long
    Dim lngUpRow As Long
    With vsAdvice
        Select Case Control.ID
        Case conMenu_Edit_Save  '�����¼
            If checkDrug = True Then
                If SaveDrug Then
                    Call UpdateDrug
                End If
            End If
        Case conMenu_Edit_DrugAuto '�Զ���ȡ
            GetCursorPos Pt
            picAdviceFilter.Left = Pt.X + (picAdviceFilter.Width / 2): picAdviceFilter.Top = Pt.Y + 300
            picAdviceFilter.Visible = Not picAdviceFilter.Visible
            picAdviceFilter.Enabled = picAdviceFilter.Visible
            If picAdviceFilter.Visible = True Then
                cmdAdviceOK.SetFocus
            Else
                .SetFocus
            End If
        Case conMenu_Edit_NewItem '������ҩ��¼
            If .TextMatrix(.Rows - 1, col_��ҩ����) = "" Then
                .Row = .Rows - 1: .Col = COL_��ʼʱ��
                .ShowCell .Row, COL_��ʼʱ��
            Else
                .Rows = vsAdvice.Rows + 1
                .Cell(flexcpBackColor, .FixedRows, col_ҩƷ���, .Rows - 1, col_ҩƷ���) = GRD_UNEDITCELL_COLOR      '����ɫ
                .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
                .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
                .Cell(flexcpBackColor, .Rows - 1, 0, vsAdvice.Rows - 1, 0) = Red_COLOR
                .Row = vsAdvice.Rows - 1: .Col = COL_��ʼʱ��
                .ShowCell .Row, COL_��ʼʱ��
            End If
        Case conMenu_Edit_Modify '�޸���ҩ��¼
            mlngLastColor = 0
            mlngEditTag = 1
            vsAdvice.Editable = flexEDKbdMouse
            lblTime.Visible = mlngEditTag = 0
             OptTime(0).Visible = lblTime.Visible: OptTime(1).Visible = lblTime.Visible: OptTime(2).Visible = lblTime.Visible: OptTime(3).Visible = lblTime.Visible
            If vsAdvice.Col = col_��ҩ���� Then Call Get��ҩ�䷽(vsAdvice.Row)
            staThis.Panels(2).Text = "��ǰģʽΪ��" & IIF(mlngEditTag = 0, "������ҩ�嵥", "�༭��ҩ�嵥")
        Case conMenu_Edit_ItemUndo '�˳��༭
            If Val(vsAdvice.Tag) = 1 Then
                If MsgBox("��ǰ����δ�������ҩ��¼,ȷ��Ҫȡ���༭��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            mlngEditTag = 0
            .Editable = flexEDNone
            Call LoadDrug
            lblTime.Visible = mlngEditTag = 0
            OptTime(0).Visible = lblTime.Visible: OptTime(1).Visible = lblTime.Visible: OptTime(2).Visible = lblTime.Visible: OptTime(3).Visible = lblTime.Visible
            picAdviceFilter.Visible = False
            staThis.Panels(2).Text = "��ǰģʽΪ��" & IIF(mlngEditTag = 0, "������ҩ�嵥", "�༭��ҩ�嵥")
        Case conMenu_Edit_Delete 'ɾ����ҩ��¼
            Call DeteleRow
        Case conMenu_Edit_DrugGrp '��ҩ��¼һ����ҩ
            If Control.Checked = True Then
                If .TextMatrix(.Row, 0) = "��" And .TextMatrix(GetUpRow(.Row), 0) <> "��" And Val(.TextMatrix(.Row, col_���)) <> Val(.TextMatrix(.Row, COL_ID)) And Val(.TextMatrix(.Row, col_���)) <> -.Row Then
                    .TextMatrix(.Row, 0) = ""
                    lngTmp = Val(.TextMatrix(.Row, col_���))
                    .TextMatrix(.Row, col_���) = ""
                    Call SetTagһ����ҩ(lngTmp)
                Else
                    If MsgBox("Ҫ������һ����ҩ��ҩƷȫ��ȡ��Ϊ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                    
                    lngTmp = Val(.TextMatrix(.Row, col_���))
                    For i = .FixedRows To .Rows - 1
                        If lngTmp = Val(.TextMatrix(i, col_���)) Then
                            .TextMatrix(i, 0) = ""
                            .TextMatrix(i, col_���) = ""
                            .TextMatrix(i, col_�Ƿ��޸�) = "1"
                            .Cell(flexcpBackColor, i, 0, i, 0) = Red_COLOR
                        End If
                    Next
                End If
            Else
                If Not Checkһ����ҩ(vsAdvice.Row) Then Exit Sub
                lngUpRow = GetUpRow(.Row)
                If Val(.TextMatrix(lngUpRow, col_���)) = 0 Then
                    If Val(.TextMatrix(lngUpRow, COL_ID)) <> 0 Then
                        .TextMatrix(lngUpRow, col_���) = Val(.TextMatrix(lngUpRow, COL_ID))
                    Else
                        .TextMatrix(lngUpRow, col_���) = -lngUpRow
                    End If
                    .TextMatrix(lngUpRow, col_�Ƿ��޸�) = "1"
                    .Cell(flexcpBackColor, lngUpRow, 0, lngUpRow, 0) = Red_COLOR
                End If
                
                'һ����ҩ����ͬ��
                .TextMatrix(.Row, col_���) = .TextMatrix(lngUpRow, col_���)
                .TextMatrix(.Row, COL_��ʼʱ��) = .TextMatrix(lngUpRow, COL_��ʼʱ��)
                .TextMatrix(.Row, COL_�÷�) = .TextMatrix(lngUpRow, COL_�÷�)
                .TextMatrix(.Row, COL_�÷�id) = .TextMatrix(lngUpRow, COL_�÷�id)
                .TextMatrix(.Row, COL_ִ��Ƶ��) = .TextMatrix(lngUpRow, COL_ִ��Ƶ��)
                .TextMatrix(.Row, COL_Ƶ�ʼ��) = .TextMatrix(lngUpRow, COL_Ƶ�ʼ��)
                .TextMatrix(.Row, COL_�����λ) = .TextMatrix(lngUpRow, COL_�����λ)
                .TextMatrix(.Row, COL_Ƶ�ʴ���) = .TextMatrix(lngUpRow, COL_Ƶ�ʴ���)
                .TextMatrix(.Row, COL_����) = .TextMatrix(lngUpRow, COL_����)
                
                .Cell(flexcpData, .Row, COL_ִ��Ƶ��) = .TextMatrix(lngUpRow, COL_ִ��Ƶ��)
                .Cell(flexcpData, .Row, COL_�÷�) = .TextMatrix(lngUpRow, COL_�÷�)
                Call SetTagһ����ҩ(.TextMatrix(.Row, col_���))
            End If
            .Tag = "1"
            .TextMatrix(.Row, col_�Ƿ��޸�) = "1"
            .Cell(flexcpBackColor, .Row, 0, .Row, 0) = Red_COLOR
        Case conMenu_File_Exit '�˳�
            Unload Me
        End Select
    End With
End Sub

Private Function Changeһ����ҩID(ByVal lngRow As Long) As Long
    Dim i As Long
    Dim lngTmp As Long
    With vsAdvice
        If Val(.TextMatrix(lngRow, col_���)) <> 0 And (Val(.TextMatrix(.Row, col_���)) = Val(.TextMatrix(.Row, COL_ID)) Or Val(.TextMatrix(lngRow, col_���)) = -lngRow) Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, col_���)) = Val(.TextMatrix(lngRow, col_���)) And i <> lngRow Then
                    If lngTmp = 0 Then
                        lngTmp = IIF(Val(.TextMatrix(i, COL_ID)) <> 0, Val(.TextMatrix(i, COL_ID)), -i)
                    End If
                    .TextMatrix(i, col_���) = lngTmp
                End If
            Next
        End If
        Changeһ����ҩID = lngTmp
    End With
End Function


Private Sub DeteleRow()
    Dim i As Long
    Dim blnTmp As Boolean
    Dim lngTmp As Long
    Dim lng��� As Long

    With vsAdvice
        If .Row < 1 Then Exit Sub
        If Val(.TextMatrix(.Row, COL_ID)) = 0 Then
            If .TextMatrix(.Row, COL_������ĿID) <> "" Then
                If MsgBox("ȷʵҪɾ��������ҩ��¼��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    '����һ����ҩ
                    If (.TextMatrix(.Row, 0) = "��" And .TextMatrix(GetDownRow(.Row), 0) = "��") Or (.TextMatrix(.Row, 0) = "��" And .TextMatrix(GetUpRow(.Row), 0) = "��") Then
                        lng��� = Val(.TextMatrix(.Row, col_���))
                        For i = .FixedRows To .Rows - 1
                            If lng��� = Val(.TextMatrix(i, col_���)) Then
                                .TextMatrix(i, 0) = ""
                                .TextMatrix(i, col_���) = ""
                                .TextMatrix(i, col_�Ƿ��޸�) = "1"
                                .Cell(flexcpBackColor, i, 0, i, 0) = Red_COLOR
                            End If
                        Next
                    Else
                        lngTmp = Changeһ����ҩID(.Row)
                        If lngTmp = 0 Then lngTmp = Val(.TextMatrix(.Row, col_���))
                    End If
                    mlngLastColor = 0
                    .RemoveItem .Row
                    If lngTmp <> 0 Then SetTagһ����ҩ (lngTmp)
                Else
                    Exit Sub
                End If
            Else
                mlngLastColor = 0
                .RemoveItem .Row
            End If
            
        Else
            If .RowHidden(.Row) = False Then
                If MsgBox("ȷʵҪɾ��������ҩ��¼��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    '����һ����ҩ
                    If (.TextMatrix(.Row, 0) = "��" And .TextMatrix(GetDownRow(.Row), 0) = "��") Or (.TextMatrix(.Row, 0) = "��" And .TextMatrix(GetUpRow(.Row), 0) = "��") Then
                        lng��� = Val(.TextMatrix(.Row, col_���))
                        For i = .FixedRows To .Rows - 1
                            If lng��� = Val(.TextMatrix(i, col_���)) Then
                                .TextMatrix(i, 0) = ""
                                .TextMatrix(i, col_���) = ""
                                .TextMatrix(i, col_�Ƿ��޸�) = "1"
                                .Cell(flexcpBackColor, i, 0, i, 0) = Red_COLOR
                            End If
                        Next
                    Else
                        lngTmp = Changeһ����ҩID(.Row)
                        If lngTmp = 0 Then lngTmp = Val(.TextMatrix(.Row, col_���))
                    End If
                    mlngLastColor = 0
                    .RowHidden(.Row) = True
                    .TextMatrix(.Row, col_���) = ""
                    If lngTmp <> 0 Then SetTagһ����ҩ (lngTmp)
                    .TextMatrix(.Row, col_�Ƿ��޸�) = "1"
                Else
                    Exit Sub
                End If
            End If
        End If
        
        
        'Ѱ����һ������
        If .RowHidden(.Row) = True Then
            For i = .Row To 1 Step -1
                If .RowHidden(i) = False Then
                    .Row = i: .Col = col_��ҩ����: .ShowCell i, col_��ҩ����: Exit For
                End If
            Next
        End If
        
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                blnTmp = True: Exit For
            End If
        Next
        If .Rows = 1 Or (Not blnTmp) Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, col_ҩƷ���, .Rows - 1, col_ҩƷ���) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Row = .Rows - 1: .Col = COL_��ʼʱ��
            .ShowCell .Rows - 1, COL_��ʼʱ��
        End If
        .Tag = "1"
    End With
End Sub

Private Function Get��ҩ�䷽(lngRow As Long) As String
    Dim strSQL As String, strTmp
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_ID)) <> 0 And .TextMatrix(lngRow, col_ҩƷ���) = "�в�ҩ" And .TextMatrix(lngRow, col_�䷽����) = "" Then
            strSQL = "Select a.�䷽id, a.���, a.������Ŀid, a.�շ�ϸĿid, a.����, a.��ע, b.����, b.���㵥λ From ������ҩ�䷽ A, ������ĿĿ¼ B Where a.������Ŀid = b.Id And a.�䷽id =[1]  order by A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    strTmp = strTmp & "[�䷽����]" & rsTmp!���� & "<Data>" & Val(rsTmp!������ĿID) & "<Data>" & Val(rsTmp!�շ�ϸĿID) & "<Data>" & FormatEx(NVL(rsTmp!����), 5) & "<Data>" & rsTmp!��ע & "<Data>" & rsTmp!���㵥λ
                    rsTmp.MoveNext
                Next
            End If
            Get��ҩ�䷽ = strTmp
            .TextMatrix(lngRow, col_�䷽����) = strTmp
            .Cell(flexcpData, lngRow, col_�䷽����) = strTmp
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdAdviceOK_Click()
    '��ȡ������ʷҽ����¼
    Dim strSQL As String
    Dim strType As String
    Dim strTime As String
    Dim str��Դ As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, i As Long
    
    On Error GoTo errH
    strType = "(a.������� In (" & IIF(chkXY.value = 1, "'5', '6'", "") & IIF(chkZY.value = 1, IIF(chkXY.value = 1, ",", "") & "'7'", "") & ")" & IIF(chkZY.value = 1, "Or a.������� = 'E' And (c.�������� = '3'))", ")")
    strTime = " And A.��ʼִ��ʱ�� between [3] and [4]"
    If chkMZ.value = 1 And chkZY.value = 1 Then
        str��Դ = ""
    Else
        str��Դ = IIF(chkMZ.value = 1, " And a.��ҳid is null", " And A.�Һŵ� is null")
    End If
    strSQL = "Select a.Id, a.���id As ���, a.������� As ҩƷ���, a.ҽ������ As ��ҩ����, a.ҽ������ As ҽ������, a.������Ŀid, a.�շ�ϸĿid, a.����, a.��ʼִ��ʱ�� As ��ʼʱ��," & vbNewLine & _
            "       a.ִ����ֹʱ�� As ��ֹʱ��, decode(a.������Դ,1,a.�ܸ�����/e.�����װ,2,a.�ܸ�����/e.סԺ��װ,a.�ܸ�����) as �ܸ�����, a.��������, a.ִ��Ƶ��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, b.������Ŀid As ��ҩid, c.���㵥λ, b.ҽ������ As �÷�, d.���� As ��ҩ�÷�,E.סԺ��λ" & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ����¼ B, ������ĿĿ¼ C, ������ĿĿ¼ D,ҩƷ��� E" & vbNewLine & _
            "Where a.���id = b.Id And a.������Ŀid = c.Id And A.�շ�ϸĿid=E.ҩƷid(+) And b.������Ŀid = d.Id " & vbNewLine & _
            " And a.����id = [1] And (nvl(a.��ҳid,0) <> [2])" & vbNewLine & _
            " And " & strType & strTime & str��Դ & vbNewLine & _
            " and nvl(a.ҽ��״̬,0)<>4" & vbNewLine & _
            "Order By a.����id,a.��ҳid,a.�Һŵ�,a.���,a.��ʼִ��ʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, CDate(dtpStartTime.value), CDate(dtpStopTime.value))
    With vsAdvice
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             If .TextMatrix(.Rows - 1, col_��ҩ����) = "" And Val(.TextMatrix(.Rows - 1, COL_ID)) = 0 Then .Rows = .Rows - 1
             For i = 1 To rsTmp.RecordCount
                If (rsTmp!ҩƷ��� & "" = "7" Or rsTmp!ҩƷ��� & "" = "E") And Val(.Cell(flexcpData, .Rows - 1, col_���)) = Val(rsTmp!��� & "") Then
                    If rsTmp!ҩƷ��� & "" = "7" Then
                        .TextMatrix(.Rows - 1, col_�䷽����) = .TextMatrix(.Rows - 1, col_�䷽����) & "[�䷽����]" & rsTmp!��ҩ���� & "<Data>" & Val(rsTmp!������ĿID & "") & "<Data>" & Val(rsTmp!�շ�ϸĿID & "") & "<Data>" & FormatEx(NVL(rsTmp!��������), 5) & "<Data>" & rsTmp!ҽ������ & "<Data>" & rsTmp!���㵥λ
                    Else
                        .TextMatrix(.Rows - 1, col_�巨id) = Val(rsTmp!������ĿID & "")
                    End If
                Else
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                        
                    '������
                    .TextMatrix(lngRow, COL_����ID) = mlng����ID
                    .TextMatrix(lngRow, COL_��ҳID) = mlng��ҳID
                    .TextMatrix(lngRow, COL_��ҩ��Դ) = 1
                    .TextMatrix(lngRow, COL_������ĿID) = Val(rsTmp!������ĿID & "")
                    .TextMatrix(lngRow, COL_�շ�ϸĿID) = Val(rsTmp!�շ�ϸĿID & "")
                    .TextMatrix(lngRow, COL_Ƶ�ʼ��) = Val(rsTmp!Ƶ�ʼ�� & "")
                    .TextMatrix(lngRow, COL_�����λ) = rsTmp!�����λ & ""
                    .TextMatrix(lngRow, COL_�÷�id) = Val(rsTmp!��ҩid & "")
                    .TextMatrix(lngRow, COL_��ֹʱ��) = Format(rsTmp!��ֹʱ�� & "", "yyyy-mm-dd hh:mm")
                    .TextMatrix(lngRow, COL_��ʼʱ��) = Format(rsTmp!��ʼʱ�� & "", "yyyy-mm-dd hh:mm")
                    .TextMatrix(lngRow, col_ҩƷ���) = decode(rsTmp!ҩƷ��� & "", "5", "����ҩ", "6", "�г�ҩ", "�в�ҩ")
                    .TextMatrix(lngRow, col_��ҩ����) = IIF(.TextMatrix(lngRow, col_ҩƷ���) = "�в�ҩ", rsTmp!�÷� & "", rsTmp!��ҩ���� & "")
                    .TextMatrix(lngRow, COL_�÷�) = IIF(.TextMatrix(lngRow, col_ҩƷ���) = "�в�ҩ", rsTmp!��ҩ�÷� & "", rsTmp!�÷� & "")
                    .TextMatrix(lngRow, COL_��������) = IIF(.TextMatrix(lngRow, col_ҩƷ���) = "�в�ҩ", "", FormatEx(NVL(rsTmp!��������), 5))
                    .TextMatrix(lngRow, COL_�ܸ�����) = FormatEx(NVL(rsTmp!�ܸ�����), 5)
                    .TextMatrix(lngRow, COL_������λ) = IIF(.TextMatrix(lngRow, col_ҩƷ���) = "�в�ҩ", "", rsTmp!���㵥λ & "")
                    .TextMatrix(lngRow, COL_������λ) = IIF(.TextMatrix(lngRow, col_ҩƷ���) = "�в�ҩ", "��", rsTmp!סԺ��λ & "")
                    .Cell(flexcpData, lngRow, col_���) = Val(rsTmp!��� & "")
                    
                    If .TextMatrix(lngRow, col_ҩƷ���) = "�в�ҩ" Then
                        .TextMatrix(lngRow, col_���) = ""
                    Else
                        If .Cell(flexcpData, lngRow, col_���) = .Cell(flexcpData, lngRow - 1, col_���) And .Cell(flexcpData, lngRow, col_���) <> "" Then
                            If .TextMatrix(lngRow - 1, col_���) = "" Then .TextMatrix(lngRow - 1, col_���) = -(lngRow - 1)
                            .TextMatrix(lngRow, col_���) = .TextMatrix(lngRow - 1, col_���)
                        End If
                    End If

                    If rsTmp!���� & "" = "" Then
                        If rsTmp!��ֹʱ�� & "" <> "" And rsTmp!��ʼʱ�� & "" <> "" Then
                            .TextMatrix(lngRow, COL_����) = FormatEx(NVL(DateDiff("d", CDate(rsTmp!��ʼʱ�� & ""), CDate(rsTmp!��ֹʱ�� & ""))), 5)
                        End If
                    Else
                        .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsTmp!����), 5)
                    End If
                    .TextMatrix(lngRow, COL_ִ��Ƶ��) = rsTmp!ִ��Ƶ�� & ""
                    
                    If rsTmp!ҩƷ��� & "" = "7" Then
                        .TextMatrix(lngRow, col_�䷽����) = .TextMatrix(lngRow, col_�䷽����) & "[�䷽����]" & rsTmp!��ҩ���� & "<Data>" & Val(rsTmp!������ĿID & "") & "<Data>" & Val(rsTmp!�շ�ϸĿID & "") & "<Data>" & FormatEx(NVL(rsTmp!��������), 5) & "<Data>" & rsTmp!ҽ������ & "<Data>" & rsTmp!���㵥λ
                    ElseIf rsTmp!ҩƷ��� & "" = "E" Then
                        .TextMatrix(lngRow, col_�巨id) = Val(rsTmp!������ĿID & "")
                    End If
                    
                    '��������
                    .Cell(flexcpData, lngRow, COL_ִ��Ƶ��) = .TextMatrix(lngRow, COL_ִ��Ƶ��)
                    .Cell(flexcpData, lngRow, COL_�÷�) = .TextMatrix(lngRow, COL_�÷�)
                    .Cell(flexcpData, lngRow, col_��ҩ����) = .TextMatrix(lngRow, col_��ҩ����)
                    .Cell(flexcpData, lngRow, col_ҩƷ���) = decode(.TextMatrix(lngRow, col_ҩƷ���), "����ҩ", "5", "�г�ҩ", "6", "�в�ҩ", "8")
                    .Cell(flexcpBackColor, lngRow, 0, lngRow, 0) = Red_COLOR
                End If
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
             .Tag = "1"
        End If
        Call SetTagһ����ҩ
        .Cell(flexcpBackColor, .FixedRows, col_ҩƷ���, .Rows - 1, col_ҩƷ���) = GRD_UNEDITCELL_COLOR      '����ɫ
        .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
        .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
    End With
    picAdviceFilter.Visible = False
    picAdviceFilter.Enabled = picAdviceFilter.Visible
    vsAdvice.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function get��ҩ�䷽(lng��Ŀid As Long) As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Long
    
    On Error GoTo errH
    '�������䷽��Ŀ
    strSQL = "Select A.ID,A.����,b.�շ�ϸĿid as ҩƷid,A.���㵥λ,B.��������,B.ҽ������,C.���" & _
        " From ������ĿĿ¼ A,������Ŀ��� B,�շ���ĿĿ¼ C" & _
        " Where A.ID=B.������ĿID And B.�������ID=[1] And c.Id(+) = b.�շ�ϸĿid" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) And A.������� IN(1,2,3) Order By B.���"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid)
    If rsTmp.EOF Then
        MsgBox "����ҩ�䷽��ǰ����Ч���䷽��ɣ����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
        Exit Function
    Else
        For i = 1 To rsTmp.RecordCount
             strTmp = strTmp & "[�䷽����]" & rsTmp!���� & "<Data>" & Val(rsTmp!ID & "") & "<Data>" & Val(rsTmp!ҩƷID & "") & "<Data>" & FormatEx(NVL(rsTmp!��������), 5) & "<Data>" & rsTmp!ҽ������ & "<Data>" & rsTmp!���㵥λ
            rsTmp.MoveNext
        Next
        get��ҩ�䷽ = strTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cmdAdviceQuit_Click()
    If picAdviceFilter.Visible = False Then Exit Sub
    picAdviceFilter.Visible = False
    picAdviceFilter.Enabled = picAdviceFilter.Visible
    vsAdvice.SetFocus
End Sub

Private Sub Form_Activate()
    vsAdvice.SetFocus
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picFilter.Top = 520
    picFilter.Left = 10
    picFilter.Width = Me.ScaleWidth
    
    picTime.Left = picFilter.Width - picTime.Width - 50
    lblTime.Left = picTime.Left - lblTime.Width - 50
    
    vsAdvice.Left = 0
    vsAdvice.Top = picFilter.Top + picFilter.Height + 10
    vsAdvice.Width = picFilter.Width
    
    vsAdvice.Height = Me.ScaleHeight - staThis.Height - vsAdvice.Top
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Modify
        Control.Enabled = mlngEditTag = 0
    Case conMenu_Edit_ItemUndo
        Control.Enabled = mlngEditTag = 1
    Case conMenu_Edit_Save
        Control.Enabled = (mlngEditTag = 1 And Val(vsAdvice.Tag) = 1)
        Control.Visible = mlngEditTag = 1
    Case conMenu_Edit_DrugAuto
        Control.Enabled = mlngEditTag = 1
    Case conMenu_Edit_NewItem
        Control.Enabled = mlngEditTag = 1
    Case conMenu_Edit_Delete
        Control.Enabled = mlngEditTag = 1
    Case conMenu_Edit_DrugGrp
        Control.Enabled = mlngEditTag = 1
        If Control.Enabled Then
            Control.Checked = Val(vsAdvice.TextMatrix(vsAdvice.Row, col_���)) <> 0
        End If
    End Select
    If Control.ID <> conMenu_Edit_Save Then Control.Visible = Control.Enabled
End Sub

Private Sub Form_Load()

    Call InitCommandBar
    Call InitAdviceTable
    
    OptTime(0).value = True

    dtpStartTime.value = DateAdd("m", -3, Now())
    dtpStopTime.value = Now()
    
    '����ƥ��
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    '����ƥ�䷽ʽ��0-ƴ��,1-���
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ"))
    mlngEditTag = 0
    mlngLastColor = 0
    
    staThis.Panels(2).Text = "��ǰģʽΪ��" & IIF(mlngEditTag = 0, "������ҩ�嵥", "�༭��ҩ�嵥")
    Call LoadPatiInfo
    vsAdvice.Editable = flexEDNone
    Call LoadDrug
    vsAdvice.Row = vsAdvice.Rows - 1: vsAdvice.Col = COL_��ʼʱ��
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    '����win10������ ͼ����ʾ�쳣
    imgIco.Top = 45: imgIco.Left = 600: imgIco.Height = 240: imgIco.Width = 240
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlngEditTag = 1 And Val(vsAdvice.Tag) = 1 Then
        If MsgBox("��ǰ����δ�������ҩ��¼,ȷ��Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub LoadPatiInfo()
'���ܣ����ز�����Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = _
        " Select B.סԺ��,b.����,b.�Ա�,b.����,B.��Ժ����," & _
        " B.סԺҽʦ,B.��Ժ����ID,C.���� as ����,B.����,B.�������� " & _
        " From  ������ҳ B,���ű� C" & _
        " Where B.��Ժ����ID=C.ID And b.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    lblPati.Caption = "������" & rsTmp!���� & "��סԺ�ţ�" & NVL(rsTmp!סԺ��) & _
        "�����ţ�" & NVL(rsTmp!��Ժ����) & "�����ң�" & NVL(rsTmp!����) & "�����䣺" & NVL(rsTmp!����)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlcommfun.GetPubIcons

    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)

    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�༭�嵥")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ItemUndo, "ȡ���༭"): objControl.IconId = 5019
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DrugAuto, "�Զ���ȡ")
            objControl.IconId = 3587
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DrugGrp, "һ����ҩ")
            objControl.IconId = 3064
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " �˳�(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "ID;����ID;��ҳID;���;��ҩ��Դ;������ĿID;�շ�ϸĿID;Ƶ�ʼ��;�����λ;�÷�ID;�巨ID;��ֹʱ��;" & _
                "��ʼʱ��,2000,1;ҩƷ���,850,4;��ҩ����,7000,1;�÷�,2000,1;����,850,4;��λ,600,4;����,850,4;��λ,600,4;����,450,4;ִ��Ƶ��,1000,4;��ע,1000,1;�Ǽ���;�Ǽ�ʱ��;�䷽����;�Ƿ��޸�;Ƶ�ʴ���"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionFree
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightWithFocus
        .BackColorSel = &H404040


        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        .Editable = flexEDKbdMouse
        .WordWrap = True
        .AutoSize col_��ҩ����
    End With
End Sub



Public Function AdviceCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsAdvice
        If .ColHidden(lngCol) Then Exit Function
        '������������ҩ����
        If (lngCol = col_ҩƷ��� Or lngCol = COL_������λ Or lngCol = COL_�Ǽ�ʱ�� Or lngCol = COL_�Ǽ��� Or lngCol = COL_������λ) Then Exit Function
        If .TextMatrix(lngRow, col_��ҩ����) = "" Then
            If lngCol > col_��ҩ���� Then Exit Function
        End If
        If lngCol = COL_�������� And .TextMatrix(lngRow, col_ҩƷ���) = "�в�ҩ" Then Exit Function
    End With
    AdviceCellEditable = True
End Function

Private Sub EnterNextCellAdvice()
    Dim i As Long, j As Long

    With vsAdvice
        '����һ��Ԫ��ʼѭ������
        If .Row < .FixedRows Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, col_ҩƷ���, .Rows - 1, col_ҩƷ���) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 0) = Red_COLOR
            .ShowCell .Rows - 1, COL_��ʼʱ��
        End If
        For i = .Row To .Rows - 1
            For j = IIF(i = .Row, .Col + 1, COL_��ʼʱ��) To COL_��ע
                If AdviceCellEditable(i, j) Then Exit For
            Next
            If j <= COL_��ע Then Exit For
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > COL_��ע And .TextMatrix(.Rows - 1, col_��ҩ����) <> "" Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, col_ҩƷ���, .Rows - 1, col_ҩƷ���) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Cell(flexcpBackColor, .FixedRows, COL_������λ, .Rows - 1, COL_������λ) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 0) = Red_COLOR
            .ShowCell .Rows - 1, COL_��ʼʱ��
        End If
    End With
End Sub


Private Sub OptTime_Click(Index As Integer)
    LoadDrug
End Sub


Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    If vsAdvice.Editable = flexEDNone Then Exit Sub
    
    If OldRow > 0 And OldCol > 0 Then
        If OldRow < vsAdvice.Rows Then
            vsAdvice.Cell(flexcpBackColor, OldRow, OldCol, OldRow, OldCol) = mlngLastColor
        End If
    End If
    If NewRow > 0 Then
        mlngLastColor = vsAdvice.Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol)
        vsAdvice.Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = &HC0FFC0
    End If
    
    If (Not AdviceCellEditable(NewRow, NewCol)) Then
        vsAdvice.ComboList = ""
        vsAdvice.FocusRect = flexFocusLight
    Else
        vsAdvice.FocusRect = flexFocusSolid
        Select Case NewCol
            Case col_��ҩ����, COL_�÷�, COL_ִ��Ƶ��
                If NewCol = col_��ҩ���� Then Call Get��ҩ�䷽(NewRow)
                vsAdvice.ComboList = "..."
            Case Else
                vsAdvice.ComboList = ""
        End Select
    End If
End Sub


Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsAdvice
        .EditSelStart = 0
        .EditSelLength = zlcommfun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not AdviceCellEditable(Row, Col) Then
        Cancel = True
    End If
End Sub


Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    If vsAdvice.Editable = flexEDNone Then Exit Sub
    With vsAdvice
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            mblnReturn = True
            Call EnterNextCellAdvice
        Else
            If .ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsAdvice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub



Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsAdvice
        If Not KeyAscii = vbKeyReturn Then
            If Col = COL_��ʼʱ�� Or Col = COL_�Ǽ�ʱ�� Then
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            ElseIf Col = COL_�������� Or Col = COL_�ܸ����� Or Col = COL_���� Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            ElseIf Col = col_��ҩ���� And vsAdvice.TextMatrix(Row, col_ҩƷ���) = "�в�ҩ" Then
                KeyAscii = 0
            End If
            mblnReturn = False
        Else
            mblnReturn = True
        End If
    End With
End Sub



Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsAdvice.Editable = flexEDNone Then Exit Sub
    With vsAdvice
        If KeyCode = vbKeyF4 Then
            If .ComboList = "..." Then
                Call zlcommfun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Col = col_��ҩ���� Then
                DeteleRow
            ElseIf .Col = COL_�÷� Then
                .TextMatrix(.Row, COL_�÷�) = ""
                .TextMatrix(.Row, COL_�÷�id) = ""
                .Cell(flexcpData, .Row, COL_�÷�) = ""
                .TextMatrix(.Row, col_�Ƿ��޸�) = "1"
            ElseIf .Col = COL_ִ��Ƶ�� Then
                   .Cell(flexcpData, .Row, COL_ִ��Ƶ��) = ""
                    .TextMatrix(.Row, COL_Ƶ�ʼ��) = ""
                    .TextMatrix(.Row, COL_Ƶ�ʴ���) = ""
                    .TextMatrix(.Row, COL_�����λ) = ""
                    .TextMatrix(.Row, col_�Ƿ��޸�) = "1"
            End If
            .Tag = "1"
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsAdvice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strSeek As String, int���� As Integer
    Dim lng��Ŀid As Long
    Dim bytOK As Byte
    Dim strData As String, lng�巨ID As Long
    Dim blnAuto As Boolean, lngUpRow As Long
    
    On Error GoTo errH
   With vsAdvice
        Select Case Col
            Case col_��ҩ����
                If .TextMatrix(Row, col_ҩƷ���) = "�в�ҩ" Then
                    strData = .TextMatrix(Row, col_�䷽����)
                    lng�巨ID = .TextMatrix(Row, col_�巨id)
                    If frmDrugListEditEx.ShowEdit(Me, strData, lng�巨ID) Then
                        If strData = "" Then
                            .TextMatrix(Row, col_��ҩ����) = ""
                            .Cell(flexcpData, Row, col_��ҩ����) = ""
                            .TextMatrix(Row, COL_������ĿID) = ""
                            .TextMatrix(Row, COL_�շ�ϸĿID) = ""
                            .TextMatrix(Row, col_ҩƷ���) = ""
                            .Cell(flexcpData, Row, col_ҩƷ���) = ""
                            .TextMatrix(Row, COL_������λ) = ""
                            .TextMatrix(Row, COL_������λ) = ""
                            .TextMatrix(Row, COL_�÷�) = ""
                            .Cell(flexcpData, Row, COL_�÷�) = ""
                            .TextMatrix(Row, COL_�÷�id) = ""
                            .TextMatrix(Row, COL_ִ��Ƶ��) = ""
                            .TextMatrix(Row, COL_Ƶ�ʴ���) = ""
                            .Cell(flexcpData, Row, COL_ִ��Ƶ��) = ""
                            .TextMatrix(Row, COL_Ƶ�ʼ��) = ""
                            .TextMatrix(Row, COL_�����λ) = ""
                            .TextMatrix(Row, col_�䷽����) = ""
                            .TextMatrix(Row, col_�巨id) = ""
                            .TextMatrix(Row, COL_��������) = ""
                            Exit Sub
                        Else
                            .TextMatrix(Row, col_��ҩ����) = Set��ҩ�䷽(Row, strData, lng�巨ID)
                            .Cell(flexcpData, Row, col_��ҩ����) = .TextMatrix(Row, col_��ҩ����)
                        End If
                        .Tag = "1"
                        .TextMatrix(Row, col_�Ƿ��޸�) = "1"
                        .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
                    Else
                        Exit Sub
                    End If
                Else
                    Set rsTmp = frmDrugSelect.ShowSelect(Me, bytOK)
                    If bytOK = 1 And (Not rsTmp Is Nothing) Then
                        If rsTmp!��� & "" = "�䷽" Or rsTmp!��� & "" = "�в�ҩ" Then
                            If Val(.TextMatrix(Row, col_���)) <> 0 Then
                                MsgBox "����һ����ҩ��ҩƷ���붼Ϊ����ҩ���г�ҩ��", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If rsTmp!��� & "" = "�䷽" Then
                                strData = get��ҩ�䷽(Val(rsTmp!������ĿID & ""))
                            Else
                                strData = "[�䷽����]" & rsTmp!���� & "<Data>" & Val(rsTmp!������ĿID & "") & "<Data>" & Val(rsTmp!�շ�ϸĿID & "") & "<Data>0<Data><Data>" & rsTmp!���㵥λ
                            End If
                            
                            If frmDrugListEditEx.ShowEdit(Me, strData, lng�巨ID) Then
                                If strData = "" Then
                                    Exit Sub
                                Else
                                    .TextMatrix(Row, col_��ҩ����) = Set��ҩ�䷽(Row, strData, lng�巨ID)
                                    .Cell(flexcpData, Row, col_��ҩ����) = .TextMatrix(Row, col_��ҩ����)
                                    .TextMatrix(Row, COL_������ĿID) = ""
                                    .TextMatrix(Row, COL_�շ�ϸĿID) = ""
                                End If
                            Else
                                Exit Sub
                            End If
                        Else
                        
                            '�¼������Զ�һ����ҩ
                            If Val(.TextMatrix(Row, col_���)) = 0 And Val(.TextMatrix(GetUpRow(Row), col_���)) <> 0 And .Cell(flexcpData, Row, col_��ҩ����) = "" And Val(.TextMatrix(Row, COL_ID)) = 0 Then
                                .TextMatrix(Row, col_���) = Val(.TextMatrix(GetUpRow(Row), col_���))
                                Call SetTagһ����ҩ(Val(.TextMatrix(Row, col_���)))
                                blnAuto = True
                            End If
                        
                            .TextMatrix(Row, col_��ҩ����) = rsTmp!���� & IIF(rsTmp!��� & "" = "", "", "(" & rsTmp!��� & ")")
                            .Cell(flexcpData, Row, col_��ҩ����) = .TextMatrix(Row, col_��ҩ����)
                            .TextMatrix(Row, COL_������ĿID) = Val(rsTmp!������ĿID & "")
                            .TextMatrix(Row, COL_�շ�ϸĿID) = Val(rsTmp!�շ�ϸĿID & "")
                            .TextMatrix(Row, col_�巨id) = ""
                            .TextMatrix(Row, col_�䷽����) = ""
                        End If

                        .TextMatrix(Row, col_ҩƷ���) = IIF(rsTmp!��� & "" = "�䷽", "�в�ҩ", rsTmp!��� & "")
                        .Cell(flexcpData, Row, col_ҩƷ���) = decode(.TextMatrix(Row, col_ҩƷ���), "����ҩ", "5", "�г�ҩ", "6", "�в�ҩ", "8")
                        .TextMatrix(Row, COL_������λ) = IIF(.TextMatrix(Row, col_ҩƷ���) = "�в�ҩ", "", rsTmp!���㵥λ & "")
                        .TextMatrix(Row, COL_������λ) = IIF(.TextMatrix(Row, col_ҩƷ���) = "�в�ҩ", "��", rsTmp!������λ & "")
                        .TextMatrix(Row, COL_��������) = IIF(.TextMatrix(Row, col_ҩƷ���) = "�в�ҩ", "", .TextMatrix(Row, COL_��������))
                        If Val(.TextMatrix(Row, col_���)) = 0 Then
                            .TextMatrix(Row, COL_�÷�) = ""
                            .Cell(flexcpData, Row, COL_�÷�) = ""
                            .TextMatrix(Row, COL_�÷�id) = ""
                            .TextMatrix(Row, COL_ִ��Ƶ��) = ""
                            .TextMatrix(.Row, COL_Ƶ�ʴ���) = ""
                            .Cell(flexcpData, Row, COL_ִ��Ƶ��) = ""
                            .TextMatrix(Row, COL_Ƶ�ʼ��) = ""
                            .TextMatrix(Row, COL_�����λ) = ""
                            .TextMatrix(Row, COL_����) = ""
                        ElseIf blnAuto Then
                            '�Զ�һ����ҩ��ͬ������
                            lngUpRow = GetUpRow(Row)
                            .TextMatrix(Row, COL_��ʼʱ��) = .TextMatrix(lngUpRow, COL_��ʼʱ��)
                            .TextMatrix(Row, COL_�÷�) = .TextMatrix(lngUpRow, COL_�÷�)
                            .Cell(flexcpData, Row, COL_�÷�) = .Cell(flexcpData, lngUpRow, COL_�÷�)
                            .TextMatrix(Row, COL_�÷�id) = .TextMatrix(lngUpRow, COL_�÷�id)
                            .TextMatrix(Row, COL_ִ��Ƶ��) = .TextMatrix(lngUpRow, COL_ִ��Ƶ��)
                            .TextMatrix(.Row, COL_Ƶ�ʴ���) = .TextMatrix(lngUpRow, COL_Ƶ�ʴ���)
                            .Cell(flexcpData, Row, COL_ִ��Ƶ��) = .Cell(flexcpData, lngUpRow, COL_ִ��Ƶ��)
                            .TextMatrix(Row, COL_Ƶ�ʼ��) = .TextMatrix(lngUpRow, COL_Ƶ�ʼ��)
                            .TextMatrix(Row, COL_�����λ) = .TextMatrix(lngUpRow, COL_�����λ)
                            .TextMatrix(Row, COL_����) = .TextMatrix(lngUpRow, COL_����)
                        End If
                        
                        .Tag = "1"
                        .TextMatrix(Row, col_�Ƿ��޸�) = "1"
                        .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
                    End If
                End If
            Case COL_�÷�
                int���� = IIF(.TextMatrix(Row, col_ҩƷ���) = "�в�ҩ", 4, 2)
                
                lng��Ŀid = Val(.TextMatrix(.Row, COL_������ĿID))
                If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
                    If Val(.TextMatrix(.Row, COL_�շ�ϸĿID)) = 0 Then
                        strSQL = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[2] And ����>0)" & _
                            " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                            " Where A.�÷�ID=B.ID And B.������� IN(1,2,3) And A.��ĿID=[2] And A.����>0)<=1)"
                    Else
                        lng��Ŀid = Val(.TextMatrix(.Row, COL_�շ�ϸĿID))
                        strSQL = " And (A.ID IN(Select �÷�ID From ҩƷ�÷����� Where ҩƷID=[2] And ����=1)" & _
                            " Or (Select Count(A.�÷�ID) From ҩƷ�÷����� A,������ĿĿ¼ B" & _
                            " Where A.�÷�ID=B.ID And B.������� IN(1,2,3) And A.ҩƷID=[2] And A.����=1)<=1)"
                    End If
                End If
                strSQL = "Select Distinct A.ID,A.����,A.����,C.���� as ����,A.ִ�з��� as ִ�з���ID" & _
                    " From ������ĿĿ¼ A,���Ʒ���Ŀ¼ C" & _
                    " Where A.����ID=C.ID(+) And A.���='E' And A.��������=[1] And A.������� IN(1,2,3)" & strSQL & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.����"
                 vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҩ;��", False, strSeek, "", False, False, True, _
                    vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, CStr(int����), lng��Ŀid)
                    
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û�п��õĸ�ҩ;�������ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                    End If
                    Exit Sub
                Else
                    .TextMatrix(Row, COL_�÷�) = rsTmp!���� & ""
                    .TextMatrix(Row, COL_�÷�id) = Val(rsTmp!ID & "")
                    .Cell(flexcpData, Row, COL_�÷�) = .TextMatrix(Row, COL_�÷�)
                    .TextMatrix(Row, col_�Ƿ��޸�) = "1"
                    .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
                    If Val(.TextMatrix(Row, col_���)) <> 0 Then Setͬ��һ����ҩ (Row)
                    .Tag = "1"
                End If
            Case COL_ִ��Ƶ��
                strSQL = _
                    " Select Rownum as ID,A.����,A.����,A.����," & _
                    " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.���÷�Χ as ��ΧID" & _
                    " From ����Ƶ����Ŀ A" & _
                    " Where A.���÷�Χ<>[1]" & _
                    " Order by A.���÷�Χ,A.����"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ִ��Ƶ��", False, strSeek, "", False, False, True, _
                    vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, IIF(.TextMatrix(.Row, col_ҩƷ���) = "�в�ҩ", "1", "2"))
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û�п��õ�����Ƶ����Ŀ�����ȵ�ҽ��Ƶ�ʹ��������á�", vbInformation, gstrSysName
                    End If
                    Exit Sub
                Else
                    .TextMatrix(Row, COL_ִ��Ƶ��) = rsTmp!���� & ""
                    .Cell(flexcpData, Row, COL_ִ��Ƶ��) = .TextMatrix(Row, COL_ִ��Ƶ��)
                    .TextMatrix(Row, COL_Ƶ�ʼ��) = rsTmp!Ƶ�ʼ�� & ""
                    .TextMatrix(Row, COL_�����λ) = rsTmp!�����λ & ""
                    .TextMatrix(.Row, COL_Ƶ�ʴ���) = rsTmp!Ƶ�ʴ��� & ""
                    .Tag = "1"
                    .TextMatrix(Row, col_�Ƿ��޸�) = "1"
                    If Val(.TextMatrix(Row, col_���)) <> 0 Then Setͬ��һ����ҩ (Row)
                    .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
                End If
        End Select
   End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInput As String
    
    With vsAdvice
        Select Case Col
            Case COL_��ʼʱ��
                strInput = Format(zlStr.FullDate(.TextMatrix(Row, Col)), "yyyy-mm-dd hh:mm")
                If Not IsDate(strInput) Then
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Else
                    .TextMatrix(Row, Col) = strInput
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    If Val(.TextMatrix(Row, col_���)) <> 0 Then Setͬ��һ����ҩ (Row)
                End If
            Case COL_����
                If Val(.TextMatrix(Row, col_���)) <> 0 Then Setͬ��һ����ҩ (Row)
        End Select
    End With
End Sub

Private Sub vsAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'vsAdvice_ValidateEdit�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strSeek As String, int���� As Integer
    Dim lng��Ŀid As Long, strLike As String
    Dim strInput As String
    Dim lngMax As Long
    Dim strData As String
    Dim lng�巨ID As Long
    Dim lngUpRow As Long, blnAuto As Boolean

    On Error GoTo errH
   With vsAdvice
        strLike = mstrLike
        If Len(.EditText) < 2 Then strLike = "" '�Ż�
        Select Case Col
            Case col_��ҩ����
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, Row, col_��ҩ����)
                    If mblnReturn Then Call EnterNextCellAdvice
                ElseIf .EditText = .Cell(flexcpData, Row, col_��ҩ����) Then
                    If mblnReturn Then Call EnterNextCellAdvice
                Else
                    strInput = " And (A.���� Like [1] And E.����=[3]" & _
                        " Or E.���� Like [2] And E.����=[3] Or E.���� Like [2] And E.���� IN([3],3))"
                
                    If IsNumeric(.EditText) Then
                        '1X.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                        If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.���� Like [1] And E.����=[3] Or E.���� Like [2] And E.����=3)"
                    ElseIf zlcommfun.IsCharAlpha(.EditText) Then
                        'X1.����ȫ����ĸʱֻƥ�����
                        If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And E.���� Like [2] And E.����=[3]"
                    ElseIf zlcommfun.IsCharChinese(.EditText) Then
                        '��������,��ֻƥ������a
                        strInput = " And E.���� Like [2] And E.����=[3]"
                    End If
                    
                    strSQL = "Select distinct a.Id, b.Id As �շ�ϸĿid,decode(a.���,'5','����ҩ','6','�г�ҩ','7','�в�ҩ','8','�䷽') as ���, a.����, b.���, a.���㵥λ, d.ҩƷ����,C.סԺ��λ as ������λ" & _
                    " From ������ĿĿ¼ A, �շ���ĿĿ¼ B, ҩƷ��� C, ҩƷ���� D,������Ŀ���� E " & _
                    " Where c.ҩƷid= b.Id(+) And a.Id =c.ҩ��id(+) And c.ҩ��id = d.ҩ��id(+) And A.ID=E.������ĿID(+) And a.��� in ('5','6','7','8') and (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & strInput
                    
                    vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩƷĿ¼", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, UCase(.EditText) & "%", strLike & UCase(.EditText) & "%", mint���� + 1)
                    
                    If rsTmp Is Nothing Then
                        If Not blnCancel Then
                            MsgBox "δ�ҵ����õ�ҩƷ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .EditText = .Cell(flexcpData, Row, col_��ҩ����)
                        End If
                        Exit Sub
                    Else
                        If rsTmp!��� & "" = "�䷽" Or rsTmp!��� & "" = "�в�ҩ" Then
                            If Val(.TextMatrix(Row, col_���)) <> 0 Then
                                MsgBox "����һ����ҩ��ҩƷ���붼Ϊ����ҩ���г�ҩ��", vbInformation, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            If rsTmp!��� & "" = "�䷽" Then
                                strData = get��ҩ�䷽(Val(rsTmp!ID & ""))
                            Else
                                strData = "[�䷽����]" & rsTmp!���� & "<Data>" & Val(rsTmp!ID & "") & "<Data>" & Val(rsTmp!�շ�ϸĿID & "") & "<Data>0<Data><Data>" & rsTmp!���㵥λ
                            End If
                            If frmDrugListEditEx.ShowEdit(Me, strData, lng�巨ID) Then
                                If strData = "" Then
                                    .EditText = .Cell(flexcpData, Row, col_��ҩ����)
                                    Exit Sub
                                Else
                                    .EditText = Set��ҩ�䷽(Row, strData, lng�巨ID)
                                    .TextMatrix(Row, col_��ҩ����) = .EditText
                                    .Cell(flexcpData, Row, col_��ҩ����) = .EditText
                                    .TextMatrix(Row, COL_������ĿID) = ""
                                    .TextMatrix(Row, COL_�շ�ϸĿID) = ""
                                End If
                            Else
                                .EditText = .Cell(flexcpData, Row, col_��ҩ����)
                                Exit Sub
                            End If
                        Else
                            '�¼������Զ�һ����ҩ
                            If Val(.TextMatrix(Row, col_���)) = 0 And Val(.TextMatrix(GetUpRow(Row), col_���)) <> 0 And .Cell(flexcpData, Row, col_��ҩ����) = "" Then
                                .TextMatrix(Row, col_���) = Val(.TextMatrix(GetUpRow(Row), col_���))
                                Call SetTagһ����ҩ(Val(.TextMatrix(Row, col_���)))
                                blnAuto = True
                            End If
                            .EditText = rsTmp!���� & IIF(rsTmp!��� & "" = "", "", "(" & rsTmp!��� & ")")
                            .TextMatrix(Row, col_��ҩ����) = rsTmp!���� & IIF(rsTmp!��� & "" = "", "", "(" & rsTmp!��� & ")")
                            .Cell(flexcpData, Row, col_��ҩ����) = .TextMatrix(Row, col_��ҩ����)
                            .TextMatrix(Row, COL_������ĿID) = Val(rsTmp!ID & "")
                            .TextMatrix(Row, COL_�շ�ϸĿID) = Val(rsTmp!�շ�ϸĿID & "")
                            .TextMatrix(Row, col_�巨id) = ""
                            .TextMatrix(Row, col_�䷽����) = ""
                        End If
                        
                        .TextMatrix(Row, col_ҩƷ���) = IIF(rsTmp!��� & "" = "�䷽", "�в�ҩ", rsTmp!��� & "")
                        .Cell(flexcpData, Row, col_ҩƷ���) = decode(.TextMatrix(Row, col_ҩƷ���), "����ҩ", "5", "�г�ҩ", "6", "�в�ҩ", "8")
                        .TextMatrix(Row, COL_������λ) = IIF(.TextMatrix(Row, col_ҩƷ���) = "�в�ҩ", "", rsTmp!���㵥λ & "")
                        .TextMatrix(Row, COL_������λ) = IIF(.TextMatrix(Row, col_ҩƷ���) = "�в�ҩ", "��", rsTmp!������λ & "")
                        .TextMatrix(Row, COL_��������) = IIF(.TextMatrix(Row, col_ҩƷ���) = "�в�ҩ", "", .TextMatrix(Row, COL_��������))
                        
                        If Val(.TextMatrix(Row, col_���)) = 0 Then
                            .TextMatrix(Row, COL_�÷�) = ""
                            .Cell(flexcpData, Row, COL_�÷�) = ""
                            .TextMatrix(Row, COL_�÷�id) = ""
                            .TextMatrix(Row, COL_ִ��Ƶ��) = ""
                            .TextMatrix(.Row, COL_Ƶ�ʴ���) = ""
                            .Cell(flexcpData, Row, COL_ִ��Ƶ��) = ""
                            .TextMatrix(Row, COL_Ƶ�ʼ��) = ""
                            .TextMatrix(Row, COL_�����λ) = ""
                            .TextMatrix(Row, COL_����) = ""
                        ElseIf blnAuto Then
                            '�Զ�һ����ҩ��ͬ������
                            lngUpRow = GetUpRow(Row)
                            .TextMatrix(Row, COL_��ʼʱ��) = .TextMatrix(lngUpRow, COL_��ʼʱ��)
                            .TextMatrix(Row, COL_�÷�) = .TextMatrix(lngUpRow, COL_�÷�)
                            .Cell(flexcpData, Row, COL_�÷�) = .Cell(flexcpData, lngUpRow, COL_�÷�)
                            .TextMatrix(Row, COL_�÷�id) = .TextMatrix(lngUpRow, COL_�÷�id)
                            .TextMatrix(Row, COL_ִ��Ƶ��) = .TextMatrix(lngUpRow, COL_ִ��Ƶ��)
                            .TextMatrix(.Row, COL_Ƶ�ʴ���) = .TextMatrix(lngUpRow, COL_Ƶ�ʴ���)
                            .Cell(flexcpData, Row, COL_ִ��Ƶ��) = .Cell(flexcpData, lngUpRow, COL_ִ��Ƶ��)
                            .TextMatrix(Row, COL_Ƶ�ʼ��) = .TextMatrix(lngUpRow, COL_Ƶ�ʼ��)
                            .TextMatrix(Row, COL_�����λ) = .TextMatrix(lngUpRow, COL_�����λ)
                            .TextMatrix(Row, COL_����) = .TextMatrix(lngUpRow, COL_����)
                        End If
                    End If
                End If
                lngMax = 1000
            Case COL_�÷�
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, Row, COL_�÷�)
                    If mblnReturn Then Call EnterNextCellAdvice
                ElseIf .EditText = .Cell(flexcpData, Row, COL_�÷�) Then
                    If mblnReturn Then Call EnterNextCellAdvice
                Else
                    int���� = IIF(vsAdvice.TextMatrix(Row, col_ҩƷ���) = "�в�ҩ", 4, 2)
                    lng��Ŀid = Val(.TextMatrix(.Row, COL_������ĿID))
                    If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
                        If Val(.TextMatrix(.Row, COL_�շ�ϸĿID)) = 0 Then
                            strSQL = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[4] And ����>0)" & _
                                " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                                " Where A.�÷�ID=B.ID And B.������� IN(1,2,3) And A.��ĿID=[4] And A.����>0)<=1)"
                        Else
                            lng��Ŀid = Val(.TextMatrix(.Row, COL_�շ�ϸĿID))
                            strSQL = " And (A.ID IN(Select �÷�ID From ҩƷ�÷����� Where ҩƷID=[4] And ����=1)" & _
                                " Or (Select Count(A.�÷�ID) From ҩƷ�÷����� A,������ĿĿ¼ B" & _
                                " Where A.�÷�ID=B.ID And B.������� IN(1,3) And A.ҩƷID=[4] And A.����=1)<=1)"
                        End If
                    End If
         
                    strSQL = "Select Distinct A.ID,A.����,A.����,A.ִ�з��� as ִ�з���ID" & _
                        " From ������ĿĿ¼ A,������Ŀ���� B" & _
                        " Where A.ID=B.������ĿID" & _
                        " And A.���='E' And A.��������=[3] And A.������� IN(1,2,3)" & strSQL & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2])" & _
                        decode(mint����, 0, " And B.���� IN([5],3)", 1, " And B.���� IN([5],3)", "") & _
                        " Order by A.����"
                     vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҩ;��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, UCase(.EditText) & "%", strLike & UCase(.EditText) & "%", CStr(int����), lng��Ŀid, mint���� + 1)
                    If rsTmp Is Nothing Then
                        If Not blnCancel Then
                            MsgBox "δ�ҵ����õĸ�ҩ;�������ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .EditText = .Cell(flexcpData, Row, COL_�÷�)
                        End If
                        Exit Sub
                    Else
                        .EditText = rsTmp!���� & ""
                        .TextMatrix(Row, COL_�÷�) = rsTmp!���� & ""
                        .Cell(flexcpData, Row, COL_�÷�) = .TextMatrix(Row, COL_�÷�)
                        .TextMatrix(Row, COL_�÷�id) = Val(rsTmp!ID & "")
                        If Val(.TextMatrix(Row, col_���)) <> 0 Then Setͬ��һ����ҩ (Row)
                    End If
                End If
            Case COL_ִ��Ƶ��
                strSQL = _
                    " Select Rownum as ID,A.����,A.����,A.����," & _
                    " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.���÷�Χ as ��ΧID" & _
                    " From ����Ƶ����Ŀ A" & _
                    " Where A.���÷�Χ<>[1] And (A.���� Like [2] Or Upper(A.����) Like [3]" & _
                    " Or Upper(A.����) Like [3] Or Upper(A.Ӣ������) Like [3])" & _
                    " Order by A.���÷�Χ,A.����"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ִ��Ƶ��", False, strSeek, "", False, False, True, _
                    vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, IIF(.TextMatrix(.Row, col_ҩƷ���) = "�в�ҩ", "1", "2"), UCase(.EditText) & "%", strLike & UCase(.EditText) & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ����õ�����Ƶ����Ŀ�����ȵ�ҽ��Ƶ�ʹ��������á�", vbInformation, gstrSysName
                        Cancel = True
                    Else
                        .EditText = .Cell(flexcpData, Row, COL_ִ��Ƶ��)
                    End If
                    Exit Sub
                Else
                    .EditText = rsTmp!���� & ""
                    .TextMatrix(Row, COL_ִ��Ƶ��) = rsTmp!���� & ""
                    .Cell(flexcpData, Row, COL_ִ��Ƶ��) = .TextMatrix(Row, COL_ִ��Ƶ��)
                    .TextMatrix(Row, COL_Ƶ�ʼ��) = rsTmp!Ƶ�ʼ�� & ""
                    .TextMatrix(Row, COL_�����λ) = rsTmp!�����λ & ""
                    .TextMatrix(Row, COL_Ƶ�ʴ���) = rsTmp!Ƶ�ʴ��� & ""
                    If Val(.TextMatrix(Row, col_���)) <> 0 Then Setͬ��һ����ҩ (Row)
                End If
                lngMax = 20
            Case COL_��������
                lngMax = 10
            Case COL_�ܸ�����
                lngMax = 10
            Case COL_����
                lngMax = 10
            Case COL_��ע
                lngMax = 1000
            Case COL_��ʼʱ��
        End Select
        
        If LenB(StrConv(.EditText, vbFromUnicode)) > lngMax And lngMax <> 0 Then
            MsgBox "���ܳ���" & lngMax & "���ַ��ĳ��ȡ�", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        
        mblnReturn = False
        .TextMatrix(Row, col_�Ƿ��޸�) = "1"
        .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
        .Tag = "1"
   End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Set��ҩ�䷽(ByVal lngRow As Long, ByVal strData As String, ByVal lng�巨ID As Long) As String
    Dim strTmp As String
    Dim arrTime As Variant, arrTmp As Variant
    Dim i As Long
    With vsAdvice
        .TextMatrix(lngRow, col_�巨id) = lng�巨ID
        .TextMatrix(lngRow, col_�䷽����) = strData
        arrTime = Split(strData, "[�䷽����]")
        For i = 1 To UBound(arrTime)
            If i = 1 Then strTmp = "��ҩ�䷽:"
            arrTmp = Split(arrTime(i), "<Data>")
            strTmp = strTmp & arrTmp(0) & " " & FormatEx(NVL(arrTmp(3)), 5) & arrTmp(5) & " " & arrTmp(4) & ","
        Next
        If strTmp <> "" Then strTmp = Left(strTmp, Len(strTmp) - 1)
        Set��ҩ�䷽ = strTmp
    End With
End Function




Private Sub vsAdvice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    
    With vsAdvice
        If mstrTip <> "" Then
            If .MouseRow = Val(Split(mstrTip, "|")(0)) And .MouseCol = Val(Split(mstrTip, "|")(1)) Then
                strInfo = Split(mstrTip, "|")(2)
            End If
        End If
    End With
    Call zlcommfun.ShowTipInfo(vsAdvice.hwnd, strInfo, True, True)
End Sub

Private Sub SetTagһ����ҩ(Optional ByVal lng��� As Long)
'���ܣ���һ����ҩ��ҽ��ǰ�ӱ�־
    Dim i As Long
    Dim lngUpRow As Long

    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If lng��� = 0 Then .TextMatrix(i, 0) = ""
            If lng��� <> 0 And Val(.TextMatrix(i, col_���)) = lng��� Then .TextMatrix(i, 0) = ""
            If Val(.TextMatrix(i, col_���)) <> 0 And ((lng��� = Val(.TextMatrix(i, col_���)) And lng��� <> 0) Or lng��� = 0) And .RowHidden(i) = False Then
                lngUpRow = GetUpRow(i)
                If lngUpRow = 0 Then
                    .TextMatrix(i, 0) = "��"
                Else
                    If Val(.TextMatrix(i, col_���)) = Val(.TextMatrix(lngUpRow, col_���)) And i <> lngUpRow Then
                        If .TextMatrix(lngUpRow, 0) = "��" Then
                            .TextMatrix(lngUpRow, 0) = "��"
                        End If
                        .TextMatrix(i, 0) = "��"
                    Else
                        .TextMatrix(i, 0) = "��"
                    End If
                End If
            End If
        Next
    End With
End Sub




Private Function Checkһ����ҩ(ByVal lngRow As Long) As Boolean
    Dim lngUpRow As Long
    With vsAdvice
        lngUpRow = GetUpRow(lngRow)
        If lngUpRow = 0 Then
             MsgBox "ǰ��û�п���һ����ҩ����ҩ�С�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If .TextMatrix(lngRow, col_ҩƷ���) = "�в�ҩ" Or .TextMatrix(lngUpRow, col_ҩƷ���) = "�в�ҩ" Then
            MsgBox "��ҩ�䷽��������Ϊһ����ҩ��", vbInformation, gstrSysName
            Exit Function
        End If
        Checkһ����ҩ = True
    End With
End Function

Private Function GetUpRow(ByVal lngRow As Long) As Long
'���ܣ�ȡ��һ����Ч��
    Dim i As Long

    With vsAdvice
        lngRow = lngRow - 1
        For i = lngRow To 1 Step -1
            If .RowHidden(i) = False Then
                GetUpRow = i: Exit For
            End If
        Next
    End With
End Function

Private Function GetDownRow(ByVal lngRow As Long) As Long
'���ܣ�ȡ��һ����Ч��
    Dim i As Long

    With vsAdvice
        lngRow = lngRow + 1
        For i = lngRow To .Rows - 1
            If .RowHidden(i) = False Then
                GetDownRow = i: Exit For
            End If
        Next
    End With
End Function


Private Function Setͬ��һ����ҩ(ByVal lngRow As Long) As Long
'���ܣ�ͬ��һ��һ����ҩ������
    Dim i As Long
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngRow, col_���)) = Val(.TextMatrix(i, col_���)) And Val(.TextMatrix(i, col_���)) <> 0 And .RowHidden(i) = False Then
                .TextMatrix(i, col_���) = .TextMatrix(lngRow, col_���)
                .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                .TextMatrix(i, COL_�÷�) = .TextMatrix(lngRow, COL_�÷�)
                .TextMatrix(i, COL_�÷�id) = .TextMatrix(lngRow, COL_�÷�id)
                .TextMatrix(i, COL_ִ��Ƶ��) = .TextMatrix(lngRow, COL_ִ��Ƶ��)
                .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                
                .Cell(flexcpData, i, COL_ִ��Ƶ��) = .TextMatrix(lngRow, COL_ִ��Ƶ��)
                .Cell(flexcpData, i, COL_�÷�) = .TextMatrix(lngRow, COL_�÷�)
                .TextMatrix(i, col_�Ƿ��޸�) = "1"
                .Cell(flexcpBackColor, i, 0, i, 0) = Red_COLOR
            End If
        Next
    End With
End Function

