VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDiagItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ϱ༭"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmDiagItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgList 
      Left            =   2520
      Top             =   7440
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
            Picture         =   "frmDiagItem.frx":000C
            Key             =   "CLASS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagItem.frx":05A6
            Key             =   "ITEM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4920
      TabIndex        =   1
      Top             =   7710
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   0
      Top             =   7710
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7695
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   13150
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "������Ŀ(&0)"
      TabPicture(0)   =   "frmDiagItem.frx":09F8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraNote(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraNote(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraNote(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tvwClass"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvwList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "��Ӧ����(&1)"
      TabPicture(1)   =   "frmDiagItem.frx":0A14
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtLocate"
      Tab(1).Control(1)=   "chkSelectAll"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "optӦ�÷�Χ(0)"
      Tab(1).Control(3)=   "optӦ�÷�Χ(1)"
      Tab(1).Control(4)=   "optӦ�÷�Χ(2)"
      Tab(1).Control(5)=   "Lvw����"
      Tab(1).Control(6)=   "lblLocate"
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtLocate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73350
         TabIndex        =   47
         ToolTipText     =   "������һ����F3��س�����λ�����F4"
         Top             =   442
         Width           =   1905
      End
      Begin MSComctlLib.ListView lvwList 
         Height          =   4050
         Left            =   5880
         TabIndex        =   46
         Top             =   4440
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   7144
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   2310
         Left            =   5880
         TabIndex        =   45
         TabStop         =   0   'False
         Tag             =   "1000"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   4075
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.CheckBox chkSelectAll 
         Appearance      =   0  'Flat
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -74835
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   480
         Width           =   675
      End
      Begin VB.OptionButton optӦ�÷�Χ 
         Appearance      =   0  'Flat
         Caption         =   "Ӧ���ڵ�ǰ��Ŀ"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   -74880
         TabIndex        =   42
         Top             =   6240
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton optӦ�÷�Χ 
         Appearance      =   0  'Flat
         Caption         =   "Ӧ����ͬ����Ŀ"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   -74880
         TabIndex        =   41
         Top             =   6600
         Width           =   5700
      End
      Begin VB.OptionButton optӦ�÷�Χ 
         Appearance      =   0  'Flat
         Caption         =   "Ӧ���ڵ�ǰ����"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   -74880
         TabIndex        =   40
         Top             =   6960
         Width           =   5775
      End
      Begin VB.Frame fraNote 
         Height          =   1875
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   5400
         Width           =   5745
         Begin VB.TextBox txtStandard 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   2565
            TabIndex        =   27
            Top             =   1440
            Width           =   1065
         End
         Begin VB.TextBox txtStandard 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   2565
            TabIndex        =   33
            Top             =   285
            Width           =   1065
         End
         Begin VB.TextBox txtStandard 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   2565
            TabIndex        =   31
            Top             =   667
            Width           =   1065
         End
         Begin VB.TextBox txtStandard 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   2565
            TabIndex        =   29
            Top             =   1050
            Width           =   1065
         End
         Begin VB.CheckBox chkStandard 
            Caption         =   "ICD-10������(&1)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   705
            TabIndex        =   34
            Top             =   300
            Width           =   1980
         End
         Begin VB.CheckBox chkStandard 
            Caption         =   "�����ж�ԭ����(&2)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   705
            TabIndex        =   32
            Top             =   675
            Width           =   1980
         End
         Begin VB.CheckBox chkStandard 
            Caption         =   "������̬ѧ����(&3)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   705
            TabIndex        =   30
            Top             =   1050
            Width           =   1980
         End
         Begin VB.CheckBox chkStandard 
            Caption         =   "��ҽ��������(&4)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   705
            TabIndex        =   28
            Top             =   1440
            Width           =   1980
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   105
            Picture         =   "frmDiagItem.frx":0A30
            Top             =   45
            Width           =   480
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "��׼���룺�ü�����Ӧ�Ĺ��ʻ���ұ�׼����"
            Height          =   180
            Index           =   2
            Left            =   630
            TabIndex        =   39
            Top             =   0
            Width           =   3600
         End
         Begin VB.Label lblStandard 
            Caption         =   "��������..."
            Height          =   180
            Index           =   0
            Left            =   3690
            TabIndex        =   38
            Top             =   345
            Width           =   1950
         End
         Begin VB.Label lblStandard 
            Caption         =   "��������..."
            Height          =   180
            Index           =   1
            Left            =   3690
            TabIndex        =   37
            Top             =   727
            Width           =   1950
         End
         Begin VB.Label lblStandard 
            Caption         =   "��������..."
            Height          =   180
            Index           =   2
            Left            =   3690
            TabIndex        =   36
            Top             =   1110
            Width           =   1950
         End
         Begin VB.Label lblStandard 
            Caption         =   "��������..."
            Height          =   180
            Index           =   3
            Left            =   3690
            TabIndex        =   35
            Top             =   1500
            Width           =   1950
         End
      End
      Begin VB.Frame fraNote 
         Height          =   1215
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   3960
         Width           =   5745
         Begin VB.CommandButton cmdSelect 
            Caption         =   "&P"
            Height          =   240
            Left            =   5055
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   315
            Width           =   285
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdClass 
            Height          =   780
            Left            =   705
            TabIndex        =   24
            Top             =   300
            Width           =   4650
            _ExtentX        =   8202
            _ExtentY        =   1376
            _Version        =   393216
            Rows            =   3
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollBars      =   0
            SelectionMode   =   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   105
            Picture         =   "frmDiagItem.frx":12FA
            Top             =   45
            Width           =   480
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "�������ࣺ����ѧ�ƽ����ԣ�������ͬʱ�����������"
            Height          =   180
            Index           =   1
            Left            =   630
            TabIndex        =   25
            Top             =   0
            Width           =   4320
         End
      End
      Begin VB.Frame fraNote 
         Height          =   3285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   5745
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   4
            Left            =   1785
            TabIndex        =   13
            Top             =   1386
            Width           =   3615
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   2
            Left            =   1785
            TabIndex        =   12
            Top             =   1019
            Width           =   1215
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   1
            Left            =   1785
            TabIndex        =   11
            Top             =   652
            Width           =   3615
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   0
            Left            =   1785
            TabIndex        =   10
            Top             =   285
            Width           =   1605
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   5
            Left            =   1785
            TabIndex        =   9
            Top             =   1753
            Width           =   3615
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   6
            Left            =   1785
            TabIndex        =   8
            Top             =   2120
            Width           =   1215
         End
         Begin VB.TextBox txtItem 
            Height          =   645
            Index           =   8
            Left            =   1785
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2490
            Width           =   3645
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   3
            Left            =   3630
            TabIndex        =   6
            Top             =   1019
            Width           =   1215
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   7
            Left            =   3630
            TabIndex        =   5
            Top             =   2120
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   105
            Picture         =   "frmDiagItem.frx":1BC4
            Top             =   45
            Width           =   480
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "���������������š����ơ������ͼ�Ҫ˵����"
            Height          =   180
            Index           =   0
            Left            =   645
            TabIndex        =   21
            Top             =   0
            Width           =   3780
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "��������(&A)"
            Height          =   180
            Index           =   5
            Left            =   750
            TabIndex        =   20
            Top             =   1830
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "Ӣ������(&E)"
            Height          =   180
            Index           =   4
            Left            =   735
            TabIndex        =   19
            Top             =   1455
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "���Ƽ���(&S)              (ƴ��)               (���)"
            Height          =   180
            Index           =   2
            Left            =   750
            TabIndex        =   18
            Top             =   1095
            Width           =   4680
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "�������(&N)"
            Height          =   180
            Index           =   1
            Left            =   750
            TabIndex        =   17
            Top             =   720
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "��ϱ��(&D)"
            Height          =   180
            Index           =   0
            Left            =   750
            TabIndex        =   16
            Top             =   345
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "��������(&B)              (ƴ��)               (���)"
            Height          =   180
            Index           =   6
            Left            =   750
            TabIndex        =   15
            Top             =   2190
            Width           =   4680
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "��Ҫ˵��(&M)"
            Height          =   180
            Index           =   8
            Left            =   750
            TabIndex        =   14
            Top             =   2565
            Width           =   990
         End
      End
      Begin MSComctlLib.ListView Lvw���� 
         Height          =   5205
         Left            =   -74880
         TabIndex        =   44
         Top             =   840
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   9181
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lblLocate 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   -73800
         TabIndex        =   48
         Top             =   495
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmDiagItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer
Dim mlng����id As Long

Const con��ϱ�� As Integer = 0
Const con������� As Integer = 1
Const con����ƴ���� As Integer = 2
Const con��������� As Integer = 3
Const conӢ������ As Integer = 4
Const con�������� As Integer = 5
Const con����ƴ���� As Integer = 6
Const con��������� As Integer = 7
Const con��Ҫ˵�� As Integer = 8

Const conICD10���� As Integer = 0
Const con�������ж� As Integer = 1
Const con������̬ѧ As Integer = 2
Const con��ҽ���� As Integer = 3

Private Sub IniDept()
    Dim rsTemp As ADODB.Recordset
    
    '���ö�Ӧ����
    On Error GoTo errHandle
    gstrSql = " Select A.���� || '-' || A.���� ����, A.ID, Nvl(B.����id, 0) ����id, A.���� " & _
            " From ���ű� A, (Select ����id From ������Ͽ��� Where ���id = [1]) B " & _
            " Where A.ID = B.����id(+) And " & _
            " A.ID In (Select ����id From ��������˵�� Where �������� In ('�ٴ�', '���', '����', '����', '����', 'Ӫ��')) And " & _
            " (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By A.���� || '-' || A.���� "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ȡ�ٴ���ҽ���ಿ��", Val(Me.Tag))
    
    With rsTemp
        If .EOF Then
            MsgBox "û�������ٴ���ҽ���ಿ�ţ������Ź���", vbInformation, gstrSysName
            Exit Sub
        End If
        Me.Lvw����.ListItems.Clear
        Do While Not .EOF
            Me.Lvw����.ListItems.Add , "_" & !ID, !����, 1, 1
            Me.Lvw����.ListItems("_" & !ID).Tag = Nvl(!����)
            If !����ID > 0 Then
                Me.Lvw����.ListItems("_" & !ID).Checked = True
            End If
            .MoveNext
        Loop
    End With
    
    '����Ӧ�÷�Χ
    gstrSql = "Select ���� From ������Ϸ��� Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ȡ��������", mlng����id)
    
    If Not rsTemp.EOF Then
        optӦ�÷�Χ(1).Caption = "Ӧ���ڡ�" & rsTemp!���� & "�������µ�������Ŀ"
    End If
    
    gstrSql = "Select ���� From ������Ϸ��� Where ��� = 1 And �ϼ�id Is Null Start With ID = [1] Connect By ID = Prior �ϼ�id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ȡ��������", mlng����id)
    
    If Not rsTemp.EOF Then
        optӦ�÷�Χ(2).Caption = "Ӧ���ڡ�" & rsTemp!���� & "�����༰�ӷ����µ�������Ŀ"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkSelectAll_Click()
    Dim n As Integer
    Dim BlnSelect As Boolean
    
    If chkSelectAll.Value = 2 Then Exit Sub
    
    BlnSelect = (chkSelectAll.Value = 1)
    
    With Lvw����
        For n = 1 To .ListItems.Count
            .ListItems(n).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub chkStandard_Click(Index As Integer)
    If Me.chkStandard(Index).Value = 1 Then
        Me.txtStandard(Index).Enabled = True
        Me.txtStandard(Index).BackColor = &H80000005
        Me.txtStandard(Index).SetFocus
    Else
        Me.txtStandard(Index).Enabled = False
        Me.txtStandard(Index).BackColor = &H8000000F
    End If
End Sub

Private Sub chkStandard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim lngItemId As Long, StrClass As String, strCollate As String
    Dim strDeptId As String
    Dim n As Integer
    Dim intӦ�÷�Χ As Integer
    
    If Trim(Me.txtItem(con��ϱ��).Text) = "" Then
        MsgBox "�����������", vbExclamation, gstrSysName
        Me.txtItem(con��ϱ��).SetFocus
        Exit Sub
    End If
    If Trim(Me.txtItem(con�������).Text) = "" Then
        MsgBox "���Ʊ�������", vbExclamation, gstrSysName
        Me.txtItem(con�������).SetFocus
        Exit Sub
    End If
    For intCount = Me.txtItem.LBound To Me.txtItem.UBound
        Select Case intCount
        Case con�������, con��������, con��Ҫ˵��
            If LenB(StrConv(Trim(Me.txtItem(intCount).Text), vbFromUnicode)) > Me.txtItem(intCount).MaxLength Then
                MsgBox Me.lblItem(intCount).Caption & "����" & Me.txtItem(intCount).MaxLength & "�ĳ�������", vbExclamation, gstrSysName
                Me.txtItem(intCount).SetFocus
                Exit Sub
            End If
        End Select
    Next
    
    StrClass = ""
    With Me.hgdClass
        For intCount = 0 To .Rows - 1
            If .RowData(intCount) <> 0 Then
                StrClass = StrClass & "," & .RowData(intCount)
            End If
        Next
        If StrClass = "" Then
            MsgBox "��������һ�ּ�����Ϸ���", vbExclamation, gstrSysName
            .SetFocus
            Exit Sub
        Else
            StrClass = Mid(StrClass, 2)
        End If
    End With
    
    strCollate = ""
    For intCount = Me.chkStandard.LBound To Me.chkStandard.UBound
        If Me.chkStandard(intCount).Value = 1 Then
            If Val(Me.txtStandard(intCount).Tag) <> 0 Then
                strCollate = strCollate & "," & Me.txtStandard(intCount).Tag
            End If
        End If
    Next
    If strCollate <> "" Then
        strCollate = Mid(strCollate, 2)
    End If
    
    '��Ӧ��������
    For n = 1 To Lvw����.ListItems.Count
        If Lvw����.ListItems(n).Checked = True Then
            strDeptId = IIf(strDeptId = "", Mid(Lvw����.ListItems(n).Key, 2), strDeptId & "," & Mid(Lvw����.ListItems(n).Key, 2))
        End If
    Next
    
    For n = 0 To optӦ�÷�Χ.UBound
        If optӦ�÷�Χ(n).Value = True Then
            intӦ�÷�Χ = n
            Exit For
        End If
    Next
    
    Err = 0: On Error GoTo ErrHand
    If Me.Tag = "����" Then
        lngItemId = zlDatabase.GetNextId("�������Ŀ¼")
        gstrSql = "zl_�������Ŀ¼_Insert(" & _
            lngItemId & "," & _
            "'" & Trim(Me.txtItem(con��ϱ��).Text) & "'," & _
            "'" & Trim(Me.txtItem(con�������).Text) & "'," & _
            "" & _
            "'" & Trim(Me.txtItem(con����ƴ����).Text) & "'," & _
            "'" & Trim(Me.txtItem(con���������).Text) & "'," & _
            "'" & Trim(Me.txtItem(conӢ������).Text) & "'," & _
            "'" & Trim(Me.txtItem(con��������).Text) & "'," & _
            "'" & Trim(Me.txtItem(con����ƴ����).Text) & "'," & _
            "'" & Trim(Me.txtItem(con���������).Text) & "'," & _
            "'" & Trim(Me.txtItem(con��Ҫ˵��).Text) & "'," & _
            IIf(Me.lblNote(0).Tag = "��ҽ", 1, 2) & "," & _
            "'" & StrClass & "','" & strCollate & "'," & _
            mlng����id & ",'" & strDeptId & "'," & intӦ�÷�Χ & ")"
    Else
        lngItemId = Me.Tag
        gstrSql = "zl_�������Ŀ¼_Update(" & _
            lngItemId & "," & _
            "'" & Trim(Me.txtItem(con��ϱ��).Text) & "'," & _
            "'" & Trim(Me.txtItem(con�������).Text) & "'," & _
            "" & _
            "'" & Trim(Me.txtItem(con����ƴ����).Text) & "'," & _
            "'" & Trim(Me.txtItem(con���������).Text) & "'," & _
            "'" & Trim(Me.txtItem(conӢ������).Text) & "'," & _
            "'" & Trim(Me.txtItem(con��������).Text) & "'," & _
            "'" & Trim(Me.txtItem(con����ƴ����).Text) & "'," & _
            "'" & Trim(Me.txtItem(con���������).Text) & "'," & _
            "'" & Trim(Me.txtItem(con��Ҫ˵��).Text) & "'," & _
            IIf(Me.lblNote(0).Tag = "��ҽ", 1, 2) & "," & _
            "'" & StrClass & "','" & strCollate & "'," & _
            mlng����id & ",'" & strDeptId & "'," & intӦ�÷�Χ & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    With Me.tvwClass
        .Left = Me.fraNote(1).Left + Me.hgdClass.Left + Me.hgdClass.ColWidth(1)
        .Top = Me.fraNote(1).Top + Me.cmdSelect.Top + Me.cmdSelect.Height
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    '��Ļ�������
    If Me.lblNote(0).Tag = "��ҽ" Then
        Me.lblNote(0).Caption = "��ҽ���������������š����ơ������ͼ�Ҫ˵����"
        Me.chkStandard(3).Visible = False
        Me.txtStandard(3).Visible = False
        Me.lblStandard(3).Visible = False
'        Me.fraNote(2).Height = 1440
'        Me.cmdHelp.Top = 6420
'        Me.cmdOK.Top = 6420
'        Me.cmdCancel.Top = 6420
'        Me.Height = 7305
    Else
        Me.lblNote(0).Caption = "��ҽ���������������š����ơ������ͼ�Ҫ˵����"
        Me.chkStandard(3).Visible = True
        Me.txtStandard(3).Visible = True
        Me.lblStandard(3).Visible = True
'        Me.fraNote(2).Height = 1875
'        Me.cmdHelp.Top = 6855
'        Me.cmdOK.Top = 6855
'        Me.cmdCancel.Top = 6855
'        Me.Height = 7740
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    mlng����id = Val(Me.txtItem(0).Tag)

    '����ѡ����װ��
    gstrSql = "select ID,�ϼ�ID,����,����,����" & _
            " From ������Ϸ���" & _
            " Where ��� = [1] " & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.lblNote(0).Tag = "��ҽ", 1, 2))
    
    With rsTemp
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "CLASS")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "CLASS")
            End If
            objNode.Sorted = True
            .MoveNext
        Loop
    End With
    
    '���Ƶ���д
    gstrSql = "select ID,����,����,˵��" & _
            " From �������Ŀ¼" & _
            " Where ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "����", -1, Me.Tag))
    
    Me.txtItem(con��ϱ��).MaxLength = rsTemp.Fields("����").DefinedSize
    Me.txtItem(con�������).MaxLength = rsTemp.Fields("����").DefinedSize
    Me.txtItem(con��Ҫ˵��).MaxLength = rsTemp.Fields("˵��").DefinedSize
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Me.txtItem(con��ϱ��).Text = rsTemp!����
        Me.txtItem(con�������).Text = rsTemp!����
        Me.txtItem(con��Ҫ˵��).Text = IIf(IsNull(rsTemp!˵��), "", rsTemp!˵��)
    Else
        gstrSql = "select nvl(max(����),'000000') as ����" & _
                " From �������Ŀ¼" & _
                " Where ��� = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.lblNote(0).Tag = "��ҽ", 1, 2))
        
        Me.txtItem(con��ϱ��).Text = Right(String(10, "0") & Val(rsTemp!����) + 1, Len(rsTemp!����))
    End If
    
    '������д
    gstrSql = "select nvl(����,'') as ����, ����, nvl(����,'') as ����, ����" & _
            " From ������ϱ���" & _
            " Where ���id=[1] " & _
            " Order by ����,����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "����", -1, Me.Tag))
    
    With rsTemp
        Me.txtItem(conӢ������).MaxLength = .Fields("����").DefinedSize
        Me.txtItem(con��������).MaxLength = .Fields("����").DefinedSize
        Me.txtItem(con����ƴ����).MaxLength = .Fields("����").DefinedSize
        Me.txtItem(con���������).MaxLength = .Fields("����").DefinedSize
        Me.txtItem(con����ƴ����).MaxLength = .Fields("����").DefinedSize
        Me.txtItem(con���������).MaxLength = .Fields("����").DefinedSize
        Do While Not .EOF
            Select Case !����
            Case 1
                If !���� = 2 Then
                    Me.txtItem(con���������).Text = !����
                Else
                    Me.txtItem(con����ƴ����).Text = !����
                End If
            Case 2
                Me.txtItem(conӢ������).Text = !����
            Case 9
                Me.txtItem(con��������).Text = !����
                If !���� = 2 Then
                    Me.txtItem(con���������).Text = !����
                Else
                    Me.txtItem(con����ƴ����).Text = !����
                End If
            End Select
            .MoveNext
        Loop
    End With
    
    '����������д(��������)
    gstrSql = "select I.ID,I.����,I.����" & _
            " from ����������� R,������Ϸ��� I" & _
            " where R.����ID=I.ID and R.���id=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "����", -1, Me.Tag))
        
    With rsTemp
        Do While Not .EOF
            Me.hgdClass.RowData(.AbsolutePosition - 1) = !ID
            Me.hgdClass.TextMatrix(.AbsolutePosition - 1, 1) = .AbsolutePosition & "."
            Me.hgdClass.TextMatrix(.AbsolutePosition - 1, 2) = "[" & !���� & "]" & !����
            If .AbsolutePosition >= 3 Then Exit Do
            .MoveNext
        Loop
        
        '����������д
        gstrSql = "select distinct ��� from ��������Ŀ¼"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Activate")
'        Call SQLTest
        Do While Not rsTemp.EOF
            Select Case rsTemp!���
            Case "B"
                Me.chkStandard(con��ҽ����).Enabled = True
                Me.lblStandard(con��ҽ����).Caption = ""
            Case "D"
                Me.chkStandard(conICD10����).Enabled = True
                Me.lblStandard(conICD10����).Caption = ""
            Case "M"
                Me.chkStandard(con������̬ѧ).Enabled = True
                Me.lblStandard(con������̬ѧ).Caption = ""
            Case "Y"
                Me.chkStandard(con�������ж�).Enabled = True
                Me.lblStandard(con�������ж�).Caption = ""
            End Select
            rsTemp.MoveNext
        Loop
    End With
    
    gstrSql = "select I.���,I.ID,I.����,I.����" & _
            " from ������϶��� R,��������Ŀ¼ I" & _
            " where R.����ID=I.ID and R.���id=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "����", -1, Me.Tag))
    
    With rsTemp
        Do While Not .EOF
            Select Case !���
            Case "B"    '��ҽ��������
                Me.chkStandard(con��ҽ����).Value = 1
                Me.txtStandard(con��ҽ����).BackColor = &H80000005
                Me.txtStandard(con��ҽ����).Tag = !ID
                Me.txtStandard(con��ҽ����).Text = !����
                Me.lblStandard(con��ҽ����).Caption = TextInLength(!����, lblStandard(con��ҽ����).Width)
            Case "D"    'ICD-10��������
                Me.chkStandard(conICD10����).Value = 1
                Me.txtStandard(conICD10����).BackColor = &H80000005
                Me.txtStandard(conICD10����).Tag = !ID
                Me.txtStandard(conICD10����).Text = !����
                Me.lblStandard(conICD10����).Caption = TextInLength(!����, lblStandard(conICD10����).Width)
            Case "M"    '������̬ѧ����
                Me.chkStandard(con������̬ѧ).Value = 1
                Me.txtStandard(con������̬ѧ).BackColor = &H80000005
                Me.txtStandard(con������̬ѧ).Tag = !ID
                Me.txtStandard(con������̬ѧ).Text = !����
                Me.lblStandard(con������̬ѧ).Caption = TextInLength(!����, lblStandard(con������̬ѧ).Width)
            Case "Y"    '�����ж����ⲿԭ��
                Me.chkStandard(con�������ж�).Value = 1
                Me.txtStandard(con�������ж�).BackColor = &H80000005
                Me.txtStandard(con�������ж�).Tag = !ID
                Me.txtStandard(con�������ж�).Text = !����
                Me.lblStandard(con�������ж�).Caption = TextInLength(!����, lblStandard(con�������ж�).Width)
            End Select
            .MoveNext
        Loop
    End With
    
    Call IniDept
    
    Me.hgdClass.Row = 0
    Call hgdClass_RowColChange
    Me.txtItem(con��ϱ��).SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtLocate_KeyPress (vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    With Me.hgdClass
        .Redraw = False
        .ColAlignmentFixed(1) = 7
        .ColAlignmentFixed(2) = 4
        .ColWidth(0) = 0
        .ColWidth(1) = 600
        .ColWidth(2) = .Width - .ColWidth(1) - 15
        .Redraw = True
    End With
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "����", "����", 3000
        .Add , "����", "����", 900
    End With
    Me.lvwList.ColumnHeaders("����").Position = 1
    Me.Lvw����.MultiSelect = False
    For intCount = Me.lblStandard.LBound To Me.lblStandard.UBound
        Me.chkStandard(intCount).Value = 0
        Me.lblStandard(intCount).Caption = "(δ���ø����׼����)"
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.tvwClass.Visible Or Me.lvwList.Visible Then
        Me.tvwClass.Visible = False
        Me.lvwList.Visible = False
        Cancel = True
    End If
End Sub

Private Sub hgdClass_GotFocus()
    Me.hgdClass.RowSel = Me.hgdClass.Row
    Me.hgdClass.ColSel = Me.hgdClass.Cols - 1
End Sub

Private Sub hgdClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then Exit Sub
    With Me.hgdClass
        For intCount = .Row To .Rows - 2
            .RowData(intCount) = .RowData(intCount + 1)
            .TextMatrix(intCount, 2) = .TextMatrix(intCount + 1, 2)
            If .TextMatrix(intCount, 2) = "" Then
                .TextMatrix(intCount, 1) = ""
            Else
                .TextMatrix(intCount, 1) = intCount + 1 & "."
            End If
        Next
        .RowData(.Rows - 1) = 0
        .TextMatrix(.Rows - 1, 1) = ""
        .TextMatrix(.Rows - 1, 2) = ""
    End With
End Sub

Private Sub hgdClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub hgdClass_RowColChange()
    With Me.cmdSelect
        .Top = Me.hgdClass.Top + Me.hgdClass.RowHeight(0) * Me.hgdClass.Row + 15
        .Left = Me.hgdClass.Left + Me.hgdClass.Width - .Width - 15
    End With
End Sub

Private Sub lblItem_Click(Index As Integer)
    Me.txtItem(Index).SetFocus
End Sub

Private Sub lvwList_DblClick()
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwList
        Me.txtStandard(.Tag).Tag = Mid(.SelectedItem.Key, 2)
        Me.txtStandard(.Tag).Text = .SelectedItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1)
        Me.lblStandard(.Tag).Caption = TextInLength(.SelectedItem.Text, lblStandard(.Tag).Width)
        Me.txtStandard(.Tag).SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        Call lvwList_DblClick
    End Select
End Sub

Private Sub lvwList_LostFocus()
    Me.lvwList.Visible = False
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    With Me.hgdClass
        .RowData(.Row) = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .TextMatrix(.Row, 1) = .Row + 1 & "."
        .TextMatrix(.Row, 2) = Me.tvwClass.SelectedItem.Text
    End With
    Me.hgdClass.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        Call tvwClass_DblClick
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If cmdSelect Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    Select Case Index
    Case con�������, con��������, con��Ҫ˵��
        Call zlCommFun.OpenIme(True)
    End Select
    Me.txtItem(Index).SelStart = 0: Me.txtItem(Index).SelLength = 100
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case con��ϱ��
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case vbKeyReturn
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Case Else
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        End Select
        KeyAscii = 0
    Case con�������, con��������
        If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case conӢ������, con����ƴ����, con���������, con����ƴ����, con���������
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case vbKeyReturn
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii = vbKeySpace Then Exit Sub
        End Select
        KeyAscii = 0
    Case con��Ҫ˵��
        If InStr("%_'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
End Sub

Private Sub txtItem_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case con�������
        Me.txtItem(con����ƴ����).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), False)
        Me.txtItem(con���������).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), True)
    Case con��������
        Me.txtItem(con����ƴ����).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), False)
        Me.txtItem(con���������).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), True)
    End Select
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    Select Case Index
    Case con�������, con��������, con��Ҫ˵��
        Call zlCommFun.OpenIme(False)
    End Select
End Sub

Private Sub txtLocate_GotFocus()
    zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long
    
    If KeyAscii = vbKeyReturn Then
        If txtLocate.Tag <> txtLocate.Text Then
            lblLocate.Tag = ""
            txtLocate.Tag = txtLocate.Text
        End If
        
        lngStart = Val("" & lblLocate.Tag) + 1
        If lngStart >= Lvw����.ListItems.Count Then lngStart = 1
    
        For i = lngStart To Lvw����.ListItems.Count
            If Lvw����.ListItems(i).Text Like "*" & txtLocate.Text & "*" Or Lvw����.ListItems(i).Tag Like "*" & UCase(txtLocate.Text) & "*" Then
                Call Lvw����.ListItems(i).EnsureVisible
                Lvw����.ListItems(i).Selected = True
                lblLocate.Tag = i
                Lvw����.SetFocus
                Exit For
            End If
        Next
    End If
End Sub

Private Sub txtStandard_GotFocus(Index As Integer)
    Me.txtStandard(Index).SelStart = 0: Me.txtStandard(Index).SelLength = 100
End Sub

Private Sub txtStandard_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("%^&*()+|=`'"":,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,����,����,����" & _
            " from ��������Ŀ¼" & _
            " where ���='" & _
            Switch(Index = conICD10����, "D", _
                   Index = con�������ж�, "Y", _
                   Index = con������̬ѧ, "M", _
                   Index = con��ҽ����, "B") & "'" & _
            "   and (���� like [1] " & _
            "       OR ���� like [2] " & _
            "       OR ���� like [2])" & _
            " and (����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Me.txtStandard(Index).Text) & "%", gstrMatch & Trim(Me.txtStandard(Index).Text) & "%")
    
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "δ�ҵ�ָ����׼��������", vbExclamation, gstrSysName
            Me.txtStandard(Index).SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txtStandard(Index).Tag = !ID
            Me.txtStandard(Index).Text = IIf(IsNull(!����), "", !����)
            Me.lblStandard(Index).Caption = TextInLength(IIf(IsNull(!����), "", !����), lblStandard(Index).Width)
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Exit Sub
        End If
        
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !ID, !����, "ITEM", "ITEM")
            objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = !����
            .MoveNext
        Loop
        With Me.lvwList
            .Tag = Index
            .ListItems(1).Selected = True
            .Left = Me.SSTab.Width - Me.lvwList.Width - 50
            .Top = Me.fraNote(2).Top + Me.txtStandard(Index).Top - .Height
            .Visible = True
            .SetFocus
        End With
    End With
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function TextInLength(ByVal strText As String, ByVal DispLen As Single) As String
    Dim iDispNum As Integer
    TextInLength = strText
    
    On Error Resume Next
    If Me.TextWidth(strText) < DispLen Then Exit Function
    iDispNum = CInt((DispLen - Me.TextWidth("...")) / Me.TextWidth(" "))
    If Me.TextWidth(MidB(strText, 1, iDispNum) & "...") > DispLen Then iDispNum = iDispNum - 1
    TextInLength = MidB(strText, 1, iDispNum)
    TextInLength = Mid(TextInLength, 1, Len(TextInLength)) & "..."
End Function
