VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHandBackPlan 
   Caption         =   "ҩƷ��ҩ�ƻ�"
   ClientHeight    =   8640
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   11760
   Icon            =   "frmHandBackPlan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   11760
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraControl 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   7800
      Width           =   11895
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   7080
         TabIndex        =   17
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   4680
         TabIndex        =   14
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   1320
         TabIndex        =   13
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   2520
         TabIndex        =   12
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ(&P)"
         Height          =   350
         Left            =   9480
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "�޸�(&M)"
         Height          =   350
         Left            =   5880
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&E)"
         Height          =   350
         Left            =   10680
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "���(&V)"
         Height          =   350
         Left            =   8280
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmHandBackPlan.frx":038A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
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
   Begin TabDlg.SSTab tabMain 
      Height          =   7815
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   "     δ��˼ƻ�(&0)     "
      TabPicture(0)   =   "frmHandBackPlan.frx":0C1E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vsfMain(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "vsfDetail(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picHsc(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "     ����˼ƻ�(&1)     "
      TabPicture(1)   =   "frmHandBackPlan.frx":0C3A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picHsc(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "vsfMain(1)"
      Tab(1).Control(2)=   "vsfDetail(1)"
      Tab(1).ControlCount=   3
      Begin VB.PictureBox picHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   0
         Left            =   100
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   5460
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2600
         Width           =   5460
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   3255
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   3000
         Width           =   6495
         _cx             =   11456
         _cy             =   5741
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlan.frx":0C56
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
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1995
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   480
         Width           =   6495
         _cx             =   11456
         _cy             =   3519
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlan.frx":0CCB
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
      Begin VB.PictureBox picHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   -74900
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   5460
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2600
         Width           =   5460
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1935
         Index           =   1
         Left            =   -74895
         TabIndex        =   10
         Top             =   480
         Width           =   6495
         _cx             =   11456
         _cy             =   3413
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlan.frx":0D40
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   3495
         Index           =   1
         Left            =   -74895
         TabIndex        =   11
         Top             =   3000
         Width           =   6495
         _cx             =   11456
         _cy             =   6165
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlan.frx":0DB5
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
End
Attribute VB_Name = "frmHandBackPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng�ⷿID As Long
Private mintUnit As String
Private mblnRefresh As Boolean

Private mstrSqlFilter As String             '���ڹ���
Private mstrBegin As String                 '��¼Ĭ�ϵĿ�ʼʱ��
Private mstrEnd As String                   '��¼Ĭ�ϵĽ���ʱ��
Private Const MStrCaption As String = "ҩƷ��ҩ�ƻ�"

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    str����ʱ�俪ʼ As String
    str����ʱ����� As String
    str���ʱ�俪ʼ As String
    str���ʱ����� As String
    lngҩƷID As Long
    lng��Ӧ��ID As Long
    str������ As String
End Type

Private SQLCondition As Type_SQLCondition

Private Enum BillType
    δ��� = 0
    ����� = 1
End Enum

'���ܣ���ϸ�б����
Private Const mconstMainHead = "���,4,500|No,1,1000|�ɱ����,7,1200|������,1,1000|��������,7,2000|�����,1,1000|�������,1,2000|ժҪ,1,3000"
Private Const mconstDetailHead = "���,4,500|��Ӧ��,1,3000|ҩƷ����,1,1000|ҩƷ����,1,2500|��Ʒ��,1,2000|���,1,2000|������,1,2000|����,1,1000|Ч��,1,1000|��λ,1,800|����,7,1000|�ɱ���,7,1000|�ɱ����,7,1000|��װ,7,0"

Private Enum �����б�
    ��� = 0
    NO = 1
    �ɱ���� = 2
    ������ = 3
    �������� = 4
    ����� = 5
    ������� = 6
    ժҪ = 7
    
    ���� = 8
End Enum

Private Enum ��ϸ�б�
    ��� = 0
    ��Ӧ�� = 1
    ҩƷ���� = 2
    ҩƷ���� = 3
    ��Ʒ�� = 4
    ��� = 5
    ������ = 6
    ���� = 7
    Ч�� = 8
    ��λ = 9
    ���� = 10
    �ɱ��� = 11
    �ɱ���� = 12
    ��װ = 13

    ���� = 14
End Enum

Private Sub GetMainDate(ByVal intType As Integer)
    '��ȡ����ҩƷ�ƻ���¼
    'intType��0-δ���;1-�����
    
    Dim rsTmp As ADODB.Recordset
    Dim strSqlCondition As String
    
    On Error GoTo errHandle
    If SQLCondition.strNO��ʼ <> "" And SQLCondition.strNO���� <> "" Then
        strSqlCondition = strSqlCondition & " And A.No >= [1] And A.No <=[2] "
    ElseIf SQLCondition.strNO��ʼ <> "" Then
        strSqlCondition = strSqlCondition & " And A.No >= [1] "
    ElseIf SQLCondition.strNO���� <> "" Then
        strSqlCondition = strSqlCondition & " And A.No <=[2] "
    End If
    
    If intType = BillType.δ��� And SQLCondition.str����ʱ�俪ʼ <> "" And SQLCondition.str����ʱ����� <> "" Then
        strSqlCondition = strSqlCondition & " And A.�������� Between [3] And [4] "
    End If
    
    If intType = BillType.����� And SQLCondition.str���ʱ�俪ʼ <> "" And SQLCondition.str���ʱ����� <> "" Then
        strSqlCondition = strSqlCondition & " And A.������� Between [5] And [6] "
    End If
     
    If SQLCondition.lngҩƷID > 0 Then
        strSqlCondition = strSqlCondition & " And A.ҩƷid=[7] "
    End If
    
    If SQLCondition.lng��Ӧ��ID > 0 Then
        strSqlCondition = strSqlCondition & " And A.��ҩ��λID + 0 =[8] "
    End If
    
    If SQLCondition.str������ <> "" Then
        strSqlCondition = strSqlCondition & " And A.����=[9] "
    End If
    
    If intType = BillType.δ��� Then
        gstrSQL = "Select A.NO, Sum(A.�ɱ����) As �ɱ����, A.������, A.��������, A.�����, A.�������, A.ժҪ " & _
                " From ҩƷ��ҩ�ƻ� A, ��Ӧ�� B " & _
                " Where A.��ҩ��λid = B.ID And ����� Is Null " & strSqlCondition & _
                " Group By A.NO, A.������, A.��������, A.�����, A.�������, A.ժҪ " & _
                " Order By NO"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                CDate(SQLCondition.str����ʱ�俪ʼ), _
                CDate(SQLCondition.str����ʱ�����), _
                CDate(SQLCondition.str���ʱ�俪ʼ), _
                CDate(SQLCondition.str���ʱ�����), _
                SQLCondition.lngҩƷID, _
                SQLCondition.lng��Ӧ��ID, _
                SQLCondition.str������)
    Else
        gstrSQL = "Select A.NO, Sum(A.�ɱ����) As �ɱ����, A.������, A.��������, A.�����, A.�������, A.ժҪ " & _
                " From ҩƷ��ҩ�ƻ� A, ��Ӧ�� B " & _
                " Where A.��ҩ��λid = B.ID And ����� Is Not Null " & strSqlCondition & _
                " Group By A.NO, A.������, A.��������, A.�����, A.�������, A.ժҪ " & _
                " Order By NO"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                CDate(SQLCondition.str����ʱ�俪ʼ), _
                CDate(SQLCondition.str����ʱ�����), _
                CDate(SQLCondition.str���ʱ�俪ʼ), _
                CDate(SQLCondition.str���ʱ�����), _
                SQLCondition.lngҩƷID, _
                SQLCondition.lng��Ӧ��ID, _
                SQLCondition.str������)
    End If
    
    vsfMain(intType).rows = 1
    vsfDetail(intType).rows = 1
    
    If rsTmp.EOF Then Exit Sub
    
    With rsTmp
        Do While Not .EOF
            vsfMain(intType).rows = vsfMain(intType).rows + 1
            
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, �����б�.���) = .AbsolutePosition
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, �����б�.NO) = !NO
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, �����б�.�ɱ����) = zlStr.FormatEx(!�ɱ����, 2, , True)
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, �����б�.������) = Nvl(!������)
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, �����б�.��������) = Format(!��������, "yyyy-mm-dd hh:mm:ss")
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, �����б�.�����) = Nvl(!�����)
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, �����б�.�������) = Format(!�������, "yyyy-mm-dd hh:mm:ss")
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, �����б�.ժҪ) = Nvl(!ժҪ)
            
            .MoveNext
        Loop
    End With
    
    Call GetDetailDate(intType, vsfMain(intType).TextMatrix(1, �����б�.NO))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDetailDate(ByVal intType As Integer, ByVal strNo As String)
    '��ȡ��ϸҩƷ�ƻ���¼
    'intType��0-δ���;1-�����
    
    Dim rsTmp As ADODB.Recordset
    Dim strSubUnit As String
    
    '��λ����װ����
    '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
    On Error GoTo errHandle
    Select Case mintUnit
    Case 1
        strSubUnit = "D.���㵥λ ��λ,1 ��װ "
    Case 2
        strSubUnit = "B.���ﵥλ ��λ,B.�����װ ��װ "
    Case 3
        strSubUnit = "B.סԺ��λ ��λ,B.סԺ��װ ��װ "
    Case 4
        strSubUnit = "B.ҩ�ⵥλ ��λ,B.ҩ���װ ��װ "
    End Select
    
    gstrSQL = "Select Distinct A.���, P.���� As ��Ӧ��, A.ҩƷid, D.���� As ҩƷ����,D.���� As ͨ����,E.���� As ��Ʒ��, " & _
        " D.���, A.ʵ������,A.Ч��, A.�ɱ���, A.�ɱ����, A.���� As ������, A.����, " & strSubUnit & _
        " From ҩƷ��ҩ�ƻ� A, ҩƷ��� B, �շ���ĿĿ¼ D, �շ���Ŀ���� E, ��Ӧ�� P " & _
        " Where A.ҩƷid = B.ҩƷid And B.ҩƷid = D.ID And A.��ҩ��λID = P.ID And B.ҩƷid = E.�շ�ϸĿid(+) And E.����(+) = 3 " & _
        " And A.NO = [1] " & _
        " Order By A.���"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��ҩ��ϸ��Ϣ", strNo)
    
    vsfDetail(intType).rows = 1
    
    If rsTmp.EOF Then Exit Sub
    
    With rsTmp
        Do While Not .EOF
            vsfDetail(intType).rows = vsfDetail(intType).rows + 1
            
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.���) = !���
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.��Ӧ��) = !��Ӧ��
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.ҩƷ����) = !ҩƷ����
            
            If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.ҩƷ����) = !ͨ����
            Else
                vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.ҩƷ����) = IIf(IsNull(!��Ʒ��), !ͨ����, !��Ʒ��)
            End If
            
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.��Ʒ��) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
            
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.���) = !���
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.��λ) = !��λ
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.����) = zlStr.FormatEx(!ʵ������ / !��װ, 2, , True)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.�ɱ���) = zlStr.FormatEx(!�ɱ��� * !��װ, 5, , True)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.�ɱ����) = zlStr.FormatEx(!�ɱ����, 2, , True)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.������) = IIf(IsNull(!������), "", !������)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.����) = IIf(IsNull(!����), "", !����)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.��װ) = !��װ
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.Ч��) = Format(IIf(IsNull(!Ч��), "", !Ч��), "yyyy-mm-dd")
            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.Ч��) <> "" Then
                '����Ϊ��Ч��
                vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.Ч��) = Format(DateAdd("D", -1, vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, ��ϸ�б�.Ч��)), "yyyy-mm-dd")
            End If
            
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowForm(FrmMain As Form, ByVal lng�ⷿID As Long, ByVal intUnit As Integer)
    mlng�ⷿID = lng�ⷿID
    mintUnit = intUnit
    
    Me.Show vbModal, FrmMain
End Sub
Private Sub IniGrid()
    Dim i As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    
    '��ʼ�����б�δ��ˣ�
    strTemp = Split(mconstMainHead, "|")
    With vsfMain(BillType.δ���)
        .Redraw = flexRDNone
        .rows = 1
        .Cols = �����б�.����
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .ColWidth(�����б�.�����) = 0
        .ColWidth(�����б�.�������) = 0
        .Redraw = flexRDDirect
    End With
    
    '��ʼ�����б�����ˣ�
    strTemp = Split(mconstMainHead, "|")
    With vsfMain(BillType.�����)
        .Redraw = flexRDNone
        .rows = 1
        .Cols = �����б�.����
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
       
        .Redraw = flexRDDirect
    End With
    
    '��ʼ��ϸ�б�δ��ˣ�
    strTemp = Split(mconstDetailHead, "|")
    With vsfDetail(BillType.δ���)
        .Redraw = flexRDNone
        .rows = 1
        .Cols = ��ϸ�б�.����
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .Redraw = flexRDDirect
    End With
    
    '��ʼ��ϸ�б�����ˣ�
    strTemp = Split(mconstDetailHead, "|")
    With vsfDetail(BillType.�����)
        .Redraw = flexRDNone
        .rows = 1
        .Cols = ��ϸ�б�.����
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmdAdd_Click()
    frmHandBackPlanModify.ShowForm Me, mlng�ⷿID, mintUnit, mblnRefresh
    
    If mblnRefresh = True Then
        Call GetMainDate(BillType.δ���)
    End If
End Sub
Private Sub cmdDel_Click()
    With vsfMain(0)
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        If MsgBox("�Ƿ�ɾ����ҩ�ƻ���[" & .TextMatrix(.Row, �����б�.NO) & "]��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            gstrSQL = "Zl_ҩƷ��ҩ�ƻ�_Delete('" & .TextMatrix(.Row, �����б�.NO) & "')"
            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            Call GetMainDate(BillType.δ���)
        End If
    End With
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
    '���ù�������
    If frmHandBackSearch.GetSearch(Me, tabMain.Tab, _
                mlng�ⷿID, _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                SQLCondition.str����ʱ�俪ʼ, _
                SQLCondition.str����ʱ�����, _
                SQLCondition.str���ʱ�俪ʼ, _
                SQLCondition.str���ʱ�����, _
                SQLCondition.lngҩƷID, _
                SQLCondition.lng��Ӧ��ID, _
                SQLCondition.str������) = True Then
        Call cmdRefresh_Click
    Else
        If SQLCondition.str����ʱ�俪ʼ = "" Or SQLCondition.str����ʱ����� = "" Then
            SQLCondition.str����ʱ�俪ʼ = mstrBegin
            SQLCondition.str����ʱ����� = mstrEnd
        End If

        If SQLCondition.str���ʱ�俪ʼ = "" Or SQLCondition.str���ʱ����� = "" Then
            SQLCondition.str���ʱ�俪ʼ = mstrBegin
            SQLCondition.str���ʱ����� = mstrEnd
        End If
    End If
End Sub
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdModify_Click()
    With vsfMain(0)
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        frmHandBackPlanModify.ShowForm Me, mlng�ⷿID, mintUnit, mblnRefresh, .TextMatrix(.Row, �����б�.NO)
        
        If mblnRefresh = True Then
            Call GetMainDate(BillType.δ���)
        End If
    End With
End Sub

Private Sub cmdPrint_Click()
    If vsfMain(tabMain.Tab).Row = 0 Then Exit Sub
    If vsfMain(tabMain.Tab).TextMatrix(vsfMain(tabMain.Tab).Row, 0) = "" Then Exit Sub
    ReportOpen gcnOracle, glngSys, "ZL1_BILL_1300_1", Me, "No=" & vsfMain(tabMain.Tab).TextMatrix(vsfMain(tabMain.Tab).Row, �����б�.NO), "��λϵ��=" & mintUnit, 1
End Sub
Private Sub cmdRefresh_Click()
     Call GetMainDate(tabMain.Tab)
End Sub
Private Sub cmdVerify_Click()
    With vsfMain(0)
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        gstrSQL = "Zl_ҩƷ��ҩ�ƻ�_Verify('" & .TextMatrix(.Row, �����б�.NO) & "','" & UserInfo.�û����� & "')"
        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
        Call GetMainDate(BillType.δ���)
    End With
End Sub

Private Sub Form_Load()
    Dim dateCurr As Date
    
    tabMain.Tab = 0
    picHsc(0).Visible = True
    vsfMain(0).Visible = True
    vsfDetail(0).Visible = True
    
    picHsc(1).Visible = False
    vsfMain(1).Visible = False
    vsfDetail(1).Visible = False
    If InStr(1, gstrprivs, ";��ӡҩƷ��ҩ�ƻ���;") > 0 Then
        cmdPrint.Visible = True
    Else
        cmdPrint.Visible = False
        cmdVerify.Left = cmdPrint.Left
        cmdDel.Left = cmdVerify.Left - cmdVerify.Width - 100
        cmdModify.Left = cmdDel.Left - cmdDel.Width - 100
        cmdAdd.Left = cmdModify.Left - cmdModify.Width - 100
    End If
    
    Call IniGrid
    
    dateCurr = Sys.Currentdate
    mstrBegin = Format(dateCurr, "YYYY-MM") & "-01 00:00:00"
    mstrEnd = Format(dateCurr, "YYYY-MM-DD") & " 23:59:59"
    SQLCondition.str����ʱ�俪ʼ = mstrBegin
    SQLCondition.str����ʱ����� = mstrEnd
    SQLCondition.str���ʱ�俪ʼ = mstrBegin
    SQLCondition.str���ʱ����� = mstrEnd
    
    Call GetMainDate(BillType.δ���)
    Call GetMainDate(BillType.�����)
    
    RestoreWinState Me, App.ProductName, MStrCaption
        
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        vsfDetail(BillType.δ���).ColWidth(��ϸ�б�.��Ʒ��) = IIf(vsfDetail(BillType.δ���).ColWidth(��ϸ�б�.��Ʒ��) = 0, 2000, vsfDetail(BillType.δ���).ColWidth(��ϸ�б�.��Ʒ��))
        vsfDetail(BillType.�����).ColWidth(��ϸ�б�.��Ʒ��) = IIf(vsfDetail(BillType.�����).ColWidth(��ϸ�б�.��Ʒ��) = 0, 2000, vsfDetail(BillType.�����).ColWidth(��ϸ�б�.��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        vsfDetail(BillType.δ���).ColWidth(��ϸ�б�.��Ʒ��) = 0
        vsfDetail(BillType.�����).ColWidth(��ϸ�б�.��Ʒ��) = 0
    End If
End Sub


Private Sub Form_Resize()
    '����λ������
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 9255 Then
            Me.Height = 9255
        End If
        
        If Me.Width < 12165 Then
            Me.Width = 12165
        End If
    End If
    
    With fraControl
        .Left = 0
        .Top = Me.ScaleHeight - fraControl.Height - IIf(staThis.Visible, staThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = 600
    End With
    
    With tabMain
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - fraControl.Height - IIf(staThis.Visible, staThis.Height, 0)
    End With
    
    With picHsc(0)
        .Height = 45
        .Left = 100
        .Width = tabMain.Width - 200
    End With
    
    With vsfMain(0)
        .Top = 480
        .Left = 100
        .Width = tabMain.Width - 200
        .Height = picHsc(0).Top - .Top
    End With
    
    With vsfDetail(0)
        .Top = picHsc(0).Top + picHsc(0).Height + 50
        .Left = vsfMain(0).Left
        .Height = tabMain.Height - .Top - 100
        .Width = vsfMain(0).Width
    End With
    
    With picHsc(1)
        .Height = 45
        .Left = 100
        .Width = tabMain.Width - 200
    End With
    
    With vsfMain(1)
        .Top = 480
        .Left = 100
        .Width = tabMain.Width - 200
        .Height = picHsc(1).Top - .Top
    End With
    
    With vsfDetail(1)
        .Top = picHsc(1).Top + picHsc(1).Height + 50
        .Left = vsfMain(1).Left
        .Height = tabMain.Height - .Top - 100
        .Width = vsfMain(1).Width
    End With

End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, MStrCaption
End Sub

Private Sub picHsc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfMain(Index).Height + y <= 500 Or vsfDetail(Index).Height - y <= 500 Then Exit Sub
        
        picHsc(Index).Top = picHsc(Index).Top + y
        vsfMain(Index).Height = vsfMain(Index).Height + y
        vsfDetail(Index).Height = vsfDetail(Index).Height - y
        vsfDetail(Index).Top = vsfDetail(Index).Top + y
        
        Me.Refresh
    End If
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If tabMain.Tab = BillType.δ��� Then
        picHsc(0).Visible = True
        vsfMain(0).Visible = True
        vsfDetail(0).Visible = True

        picHsc(1).Visible = False
        vsfMain(1).Visible = False
        vsfDetail(1).Visible = False

        cmdAdd.Enabled = True
        cmdModify.Enabled = True
        cmdDel.Enabled = True
        cmdVerify.Enabled = True
    Else
        picHsc(0).Visible = False
        vsfMain(0).Visible = False
        vsfDetail(0).Visible = False

        picHsc(1).Visible = True
        vsfMain(1).Visible = True
        vsfDetail(1).Visible = True

        cmdAdd.Enabled = False
        cmdModify.Enabled = False
        cmdDel.Enabled = False
        cmdVerify.Enabled = False
    End If
End Sub

Private Sub vsfMain_Click(Index As Integer)
    With vsfMain(Index)
        If .Row = 0 Then Exit Sub
        Call GetDetailDate(Index, .TextMatrix(.Row, �����б�.NO))
    End With
End Sub


