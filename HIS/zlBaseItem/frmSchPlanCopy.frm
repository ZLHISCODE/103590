VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSchPlanCopy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ԤԼ--ԤԼ��������"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "frmSchPlanCopy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   6735
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabPlanCopy 
      Height          =   5775
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "����ԤԼ����"
      TabPicture(0)   =   "frmSchPlanCopy.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "����ʱ��ƻ�"
      TabPicture(1)   =   "frmSchPlanCopy.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Label5"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame4 
         Caption         =   "ԴԤԼ����"
         Height          =   5055
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   2775
         Begin VSFlex8Ctl.VSFlexGrid vsfTimeProject 
            Height          =   1140
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   3840
            Width           =   2505
            _cx             =   4419
            _cy             =   2011
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         Begin VB.ListBox lstPlan 
            Columns         =   1
            Height          =   1140
            Index           =   3
            ItemData        =   "frmSchPlanCopy.frx":047A
            Left            =   120
            List            =   "frmSchPlanCopy.frx":0487
            TabIndex        =   22
            Top             =   2265
            Width           =   2505
         End
         Begin VB.ListBox lstDevice 
            Height          =   1140
            Index           =   3
            ItemData        =   "frmSchPlanCopy.frx":04A3
            Left            =   120
            List            =   "frmSchPlanCopy.frx":04B0
            TabIndex        =   21
            Top             =   585
            Width           =   2505
         End
         Begin VB.Label Label11 
            Caption         =   "Դʱ��ƻ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   3615
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "ԴԤԼ�豸"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "ԴԤԼ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1987
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ŀ��ԤԼ����"
         Height          =   5055
         Left            =   -71160
         TabIndex        =   15
         Top             =   480
         Width           =   2775
         Begin VB.ListBox lstDevice 
            Height          =   1140
            Index           =   4
            ItemData        =   "frmSchPlanCopy.frx":04CC
            Left            =   120
            List            =   "frmSchPlanCopy.frx":04D9
            TabIndex        =   17
            Top             =   600
            Width           =   2505
         End
         Begin VB.ListBox lstPlan 
            Columns         =   1
            Height          =   1110
            Index           =   4
            ItemData        =   "frmSchPlanCopy.frx":04F5
            Left            =   120
            List            =   "frmSchPlanCopy.frx":0502
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   2280
            Width           =   2505
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfTimeProject 
            Height          =   1140
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   3825
            Width           =   2505
            _cx             =   4419
            _cy             =   2011
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         Begin VB.Label Label12 
            Caption         =   "Ŀ��ʱ��ƻ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Ŀ��ԤԼ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Ŀ��ԤԼ�豸"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ԴԤԼ����"
         Height          =   5055
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2775
         Begin VB.ListBox lstDevice 
            Height          =   1860
            Index           =   1
            ItemData        =   "frmSchPlanCopy.frx":051E
            Left            =   120
            List            =   "frmSchPlanCopy.frx":052B
            TabIndex        =   11
            Top             =   630
            Width           =   2505
         End
         Begin VB.ListBox lstPlan 
            Columns         =   1
            Height          =   1950
            Index           =   1
            ItemData        =   "frmSchPlanCopy.frx":0547
            Left            =   120
            List            =   "frmSchPlanCopy.frx":0554
            Style           =   1  'Checkbox
            TabIndex        =   10
            Top             =   3000
            Width           =   2500
         End
         Begin VB.Label Label2 
            Caption         =   "ԴԤԼ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "ԴԤԼ�豸"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ŀ��ԤԼ�豸"
         Height          =   5055
         Left            =   3840
         TabIndex        =   4
         Top             =   480
         Width           =   2775
         Begin VB.ListBox lstPlan 
            Columns         =   1
            Enabled         =   0   'False
            ForeColor       =   &H00808080&
            Height          =   1860
            Index           =   2
            ItemData        =   "frmSchPlanCopy.frx":0570
            Left            =   120
            List            =   "frmSchPlanCopy.frx":057D
            TabIndex        =   6
            Top             =   3080
            Width           =   2500
         End
         Begin VB.ListBox lstDevice 
            Height          =   1950
            Index           =   2
            ItemData        =   "frmSchPlanCopy.frx":0599
            Left            =   120
            List            =   "frmSchPlanCopy.frx":05A6
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   600
            Width           =   2500
         End
         Begin VB.Label Label10 
            Caption         =   "Ŀ��ԤԼ�豸"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Ŀ��ԤԼ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   2800
            Width           =   1455
         End
      End
      Begin VB.Label Label5 
         Caption         =   "------>      ------>        ------>"
         Height          =   975
         Left            =   -71950
         TabIndex        =   14
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "------>      ------>        ------>"
         Height          =   975
         Left            =   3050
         TabIndex        =   3
         Top             =   2520
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "����"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      TabIndex        =   1
      Top             =   5940
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4200
      TabIndex        =   0
      Top             =   5940
      Width           =   1100
   End
End
Attribute VB_Name = "frmSchPlanCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���ԤԼʱ��ƻ���ID,��ʼʱ��,����ʱ��,ԤԼ����,���㷽��
Private Enum constScheduleTimeProject
    col_SchTimeProject_ID = 0
    col_SchTimeProject_��ʼʱ�� = 1
    col_SchTimeProject_����ʱ�� = 2
    col_SchTimeProject_ԤԼ���� = 3
    col_SchTimeProject_���㷽�� = 4
End Enum

Public Function zlShowMe(frmParent As Object)
'------------------------------------------------
'���ܣ���ʾ����
'������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    Call loadDevices
    Me.Show 1, frmParent
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub loadDevices()
'------------------------------------------------
'���ܣ�����ԤԼ�豸������Դ��Ŀ��ԤԼ�豸������������һ���б��
'������
'���أ���
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngDeviceID As Long
    Dim strDeviceName As String
    Dim i As Integer

    On Error GoTo err
    For i = 1 To 4
        lstDevice(i).Clear
    Next i

    strSql = "select ID,�豸���� from Ӱ��ԤԼ�豸 order by ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Ӱ��ԤԼ��������")

    While Not rsTemp.EOF
        strDeviceName = NVL(rsTemp!�豸����)
        lngDeviceID = rsTemp!ID
        
        For i = 1 To 4
            lstDevice(i).AddItem strDeviceName
            lstDevice(i).ItemData(lstDevice(i).ListCount - 1) = lngDeviceID
        Next i
        rsTemp.MoveNext
    Wend

    '���ض�Ӧ��ԤԼ����
    For i = 1 To 4
        If lstDevice(i).ListCount >= 1 Then
            lstDevice(i).ListIndex = 0
            Call loadPlans(lstDevice(i).ItemData(lstDevice(i).ListIndex), lstPlan(i))
        End If
    Next i

    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub loadPlans(lngDeviceID As Long, objControl As Object)
'------------------------------------------------
'���ܣ�����ԤԼ����������Դ��Ŀ��ԤԼ�豸��һ�������������б��
'������ lngDeviceID -- ԤԼ�豸ID
'       objControl -- ԤԼ�����б�ؼ� ListBox�ؼ�
'���أ���
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngPlanID As Long
    Dim strPlanName As String
    Dim blnAdd As Boolean
    Dim lngPlanType As Long
    
    On Error GoTo err
    
    objControl.Clear

    strSql = "select ID,��������,�������� from Ӱ��ԤԼ���� where ԤԼ�豸ID=[1] order by ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Ӱ��ԤԼ��������", lngDeviceID)

    blnAdd = False
    While Not rsTemp.EOF
        strPlanName = NVL(rsTemp!��������)
        lngPlanID = rsTemp!ID
        lngPlanType = NVL(rsTemp!��������, 0)
        
        objControl.AddItem strPlanName
        objControl.ItemData(objControl.ListCount - 1) = lngPlanID

        rsTemp.MoveNext
    Wend

    If objControl.ListCount >= 1 Then
        objControl.ListIndex = 0
    End If
    
    If objControl.Index = 3 Or objControl.Index = 4 Then
        If objControl.ListCount >= 1 Then
            Call loadTimeProjects(CLng(objControl.ItemData(objControl.ListIndex)), vsfTimeProject(objControl.Index))
        Else
            Call loadTimeProjects(0, vsfTimeProject(objControl.Index))
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    
    If tabPlanCopy.Tab = 0 Then '���Ʒ���
        Call CopyPlan
        'ˢ�µ�ǰ��Ŀ�귽���б�
        Call loadPlans(CLng(lstDevice(2).ItemData(lstDevice(2).ListIndex)), lstPlan(2))
    Else    '����ʱ��ƻ�
        Call CopyTime
        'ˢ�µ�ǰʱ��ƻ��б�
        Call loadTimeProjects(lstPlan(4).ItemData(lstPlan(4).ListIndex), vsfTimeProject(4))
    End If
End Sub

Private Sub Form_Load()
    tabPlanCopy.Tab = 0
End Sub

Private Sub lstDevice_Click(Index As Integer)
    If lstDevice(Index).ListCount >= 1 Then
        Call loadPlans(CLng(lstDevice(Index).ItemData(lstDevice(Index).ListIndex)), lstPlan(Index))
    End If
End Sub

Private Sub lstDevice_ItemCheck(Index As Integer, Item As Integer)
    'Ŀ���豸�����ܺ�Դ�豸��ͬ
    If Index = 2 Then
        If lstDevice(2).Selected(Item) = True Then
            If lstDevice(1).List(lstDevice(1).ListIndex) = lstDevice(2).List(Item) Then
                lstDevice(2).Selected(Item) = False
            End If
        End If
    End If
End Sub

Private Sub lstDevice_KeyPress(Index As Integer, KeyAscii As Integer)
    Call zlControl.CboSetIndex(lstDevice(Index).hwnd, zlControl.CboMatchIndex(lstDevice(Index).hwnd, KeyAscii))
    
    If KeyAscii = vbKeyReturn Then
        Call lstDevice_Click(Index)
    End If
End Sub

Private Sub lstPlan_Click(Index As Integer)
    Dim i As Integer
    
    If Index = 3 Or Index = 4 Then
        If lstPlan(Index).ListCount >= 1 Then
            Call loadTimeProjects(CLng(lstPlan(Index).ItemData(lstPlan(Index).ListIndex)), vsfTimeProject(Index))
        End If
    End If
    
    If Index = 3 Then
        If lstDevice(3).ListIndex = lstDevice(4).ListIndex Then
            For i = 0 To lstPlan(4).ListCount - 1
                If lstPlan(4).List(i) = lstPlan(3).List(lstPlan(3).ListIndex) Then
                    lstPlan(4).Selected(i) = False
                    Exit For
                End If
            Next i
        End If
    End If
End Sub

Private Sub lstPlan_ItemCheck(Index As Integer, Item As Integer)
    Dim i As Integer
    
    'ѡ����Դ������Ҫȡ����Դ�豸��ͬ��Ŀ���豸
    If Index = 1 Then
        For i = 0 To lstDevice(2).ListCount - 1
            If lstDevice(2).List(i) = lstDevice(1).List(lstDevice(1).ListIndex) Then
                lstDevice(2).Selected(i) = False
                Exit For
            End If
        Next i
    End If
    If Index = 4 Then
        If lstPlan(4).Selected(Item) = True Then
            If lstDevice(3).ListIndex = lstDevice(4).ListIndex Then
                If lstPlan(3).List(lstPlan(3).ListIndex) = lstPlan(4).List(Item) Then
                    lstPlan(4).Selected(Item) = False
                End If
            End If
        End If
    End If
End Sub

Private Function CopyPlan() As Boolean
'------------------------------------------------
'���ܣ����ݽ����ϵ�ѡ�񣬸���ԤԼ����
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim strSql As String
    Dim blnInTrans As Boolean       '�Ƿ���������֮��
    Dim arrSQL() As Variant
    Dim lngZhouCount As Long
    Dim strType As String
    Dim blnReplace As Boolean
    
    On Error GoTo err
    
    '���ȼ�����ݵ���Ч��
    If lstPlan(1).SelCount = 0 Then
        MsgBox "����ѡ��ԴԤԼ���������ٵ��������ơ���ť���Ʒ�����", vbOKOnly, "ԤԼ��������--��ʾ"
        Exit Function
    End If
    
    If lstDevice(2).SelCount = 0 Then
        MsgBox "����ѡ��Ŀ��ԤԼ�豸�����ٵ��������ơ���ť���Ʒ�����", vbOKOnly, "ԤԼ��������--��ʾ"
        Exit Function
    End If
    
    '���Ʒ�����ʱ�򣬲�����ͻ��ֱ����ʾ�û�����ͬ���͵ķ�������ȫ�����滻
    If MsgBox("���Ʒ�����ʱ����ͬ���͵ķ����ᱻ�·����滻����ȷ���Ƿ���Ҫ����?", vbOKCancel, "ԤԼ��������--��ʾ") = vbCancel Then
        Exit Function
    End If
    
    arrSQL = Array()
    lngZhouCount = 0     '�Ƿ��һ���ܷ���
    For i = 0 To lstPlan(1).ListCount - 1
        If lstPlan(1).Selected(i) = True Then
            strType = Mid(lstPlan(1).List(i), 2, 1)
            If strType = "��" Then
                lngZhouCount = lngZhouCount + 1
            End If
            blnReplace = IIF(strType = "��", IIF(lngZhouCount = 1, True, False), True)
            
            For j = 0 To lstDevice(2).ListCount - 1
                If lstDevice(2).Selected(j) = True Then
                    
                    strSql = "Zl_Ӱ��ԤԼ����_����(" & lstPlan(1).ItemData(i) & "," & lstDevice(2).ItemData(j) & "," & IIF(blnReplace = True, 1, 0) & ")"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSql
                End If
            Next j
        End If
    Next i
    
    gcnOracle.BeginTrans        '��ʼ���Ʒ���
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����ԤԼ����")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    CopyPlan = True
    
    Exit Function
err:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function

Private Sub CopyTime()
'------------------------------------------------
'���ܣ����ݽ����ϵ�ѡ�񣬸���ԤԼʱ��ƻ�
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim strSql As String
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    
    On Error GoTo err
    
    '���ȼ�����ݵ���Ч��
    If lstPlan(3).ListIndex < 0 Then
        MsgBox "����ѡ��ԴԤԼ���������ٵ��������ơ���ť����ʱ��ƻ���", vbOKOnly, "ԤԼʱ��ƻ�����--��ʾ"
        Exit Sub
    End If
    
    If lstPlan(4).SelCount = 0 Then
        MsgBox "����ѡ��Ŀ��ԤԼ���������ٵ��������ơ���ť����ʱ��ƻ���", vbOKOnly, "ԤԼʱ��ƻ�����--��ʾ"
        Exit Sub
    End If
    
    '���Ƽƻ���ʱ�򣬲�����ͻ��ֱ����ʾ�û�����ͬ���͵�ʱ��ƻ�����ȫ�����滻
    If MsgBox("����ʱ��ƻ���ʱ�򣬷���������ʱ��ƻ����ᱻ�滻����ȷ���Ƿ���Ҫ����?", vbOKCancel, "ԤԼʱ��ƻ�����--��ʾ") = vbCancel Then
        Exit Sub
    End If
    
    arrSQL = Array()
    For i = 0 To lstPlan(4).ListCount - 1
        If lstPlan(4).Selected(i) = True Then
            strSql = "Zl_Ӱ��ԤԼʱ��ƻ�_����(" & lstPlan(3).ItemData(lstPlan(3).ListIndex) & "," & lstPlan(4).ItemData(i) & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSql
        End If
    Next i
    
    gcnOracle.BeginTrans        '��ʼ���Ʒ���
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����ԤԼʱ��ƻ�")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    Exit Sub
err:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub loadTimeProjects(lngPlanID As Long, objGrid As VSFlexGrid)
'------------------------------------------------
'���ܣ�����ԤԼʱ��ƻ�������Դ��Ŀ��������б�
'������ lngPlanID -- ԤԼ����ID
'       objGrid -- ԤԼʱ��ƻ��б� vsfFlexGrid�ؼ�
'���أ���
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    With objGrid
        .Rows = 1
        .Cols = 5
        .FixedRows = 1
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarBoth
        .ExtendLastCol = True
        .ColWidthMax = 800
        
        '��ʾ����
        .TextMatrix(0, col_SchTimeProject_��ʼʱ��) = "��ʼʱ��"
        .TextMatrix(0, col_SchTimeProject_����ʱ��) = "����ʱ��"
        .TextMatrix(0, col_SchTimeProject_ԤԼ����) = "����"
        .TextMatrix(0, col_SchTimeProject_���㷽��) = "���㷽��"
        
        .ColWidth(col_SchTimeProject_ԤԼ����) = 450
        '���غ�̨����
        .ColHidden(col_SchTimeProject_ID) = True
    End With
    
    strSql = "select ID,��ʼʱ��,����ʱ��,ԤԼ����,���㷽�� from Ӱ��ԤԼʱ��ƻ� where ԤԼ����ID =[1] Order by ��ʼʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡԤԼʱ��ƻ�", lngPlanID)
    
    If rsTemp.EOF = False Then
    
        With objGrid
            .Rows = rsTemp.RecordCount + 1
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
            
            '�����ݿ��������
            For i = 1 To rsTemp.RecordCount
                .TextMatrix(i, col_SchTimeProject_ID) = rsTemp!ID
                .TextMatrix(i, col_SchTimeProject_��ʼʱ��) = Format(NVL(rsTemp!��ʼʱ��), "HH:SS")
                .TextMatrix(i, col_SchTimeProject_����ʱ��) = Format(NVL(rsTemp!����ʱ��), "HH:SS")
                .TextMatrix(i, col_SchTimeProject_ԤԼ����) = NVL(rsTemp!ԤԼ����)
                .TextMatrix(i, col_SchTimeProject_���㷽��) = IIF(NVL(rsTemp!���㷽��) = 1, "���˴�ƽ��", "��Ŀ�ۼ�")
                rsTemp.MoveNext
            Next i
            
        End With
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub
