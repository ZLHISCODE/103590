VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7635
      TabIndex        =   50
      Top             =   885
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7635
      TabIndex        =   49
      Top             =   450
      Width           =   1100
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   6495
      Left            =   75
      TabIndex        =   52
      Top             =   105
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frm��������.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra��Χ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkType(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkType(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkType(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkType(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkType(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "�߼�(&1)"
      TabPicture(1)   =   "frm��������.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra"
      Tab(1).ControlCount=   1
      Begin VB.CheckBox chkType 
         Caption         =   "�豸(&S)"
         Height          =   195
         Index           =   2
         Left            =   3060
         TabIndex        =   24
         Tag             =   "4"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkType 
         Caption         =   "����(&M)"
         Height          =   195
         Index           =   1
         Left            =   1845
         TabIndex        =   23
         Tag             =   "2"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkType 
         Caption         =   "ҩƷ(&D)"
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   22
         Tag             =   "1"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkType 
         Caption         =   "����(&Q)"
         Height          =   195
         Index           =   3
         Left            =   4245
         TabIndex        =   25
         Tag             =   "4"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkType 
         Caption         =   "����(&W)"
         Height          =   195
         Index           =   4
         Left            =   5460
         TabIndex        =   26
         Tag             =   "4"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.Frame fra 
         Caption         =   "��������"
         Height          =   4785
         Left            =   -74916
         TabIndex        =   54
         Top             =   420
         Width           =   7215
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   11
            Left            =   960
            TabIndex        =   57
            Tag             =   "��ʼ�������"
            Top             =   3480
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   12
            Left            =   4125
            TabIndex        =   56
            Tag             =   "�����������"
            Top             =   3480
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   10
            Left            =   960
            TabIndex        =   48
            Tag             =   "�����"
            Top             =   4260
            Width           =   5985
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   9
            Left            =   960
            TabIndex        =   46
            Tag             =   "������"
            Top             =   3870
            Width           =   5985
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   8
            Left            =   4125
            TabIndex        =   44
            Tag             =   "������Ʊ��"
            Top             =   2328
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   7
            Left            =   960
            TabIndex        =   42
            Tag             =   "��ʼ��Ʊ��"
            Top             =   2328
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   4125
            TabIndex        =   40
            Tag             =   "������ⵥ�ݺ�"
            Top             =   1944
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   5
            Left            =   960
            TabIndex        =   38
            Tag             =   "��ʼ��ⵥ�ݺ�"
            Top             =   1944
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   4125
            TabIndex        =   36
            Tag             =   "��������"
            Top             =   1545
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   960
            TabIndex        =   34
            Tag             =   "��ʼ����"
            Top             =   1545
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   960
            TabIndex        =   32
            Tag             =   "����"
            Top             =   1140
            Width           =   5985
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   30
            Tag             =   "���"
            Top             =   780
            Width           =   5985
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   960
            TabIndex        =   28
            Tag             =   "Ʒ��"
            Top             =   396
            Width           =   5985
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�������"
            Height          =   180
            Index           =   11
            Left            =   192
            TabIndex        =   59
            Top             =   3555
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   3840
            TabIndex        =   58
            Top             =   3540
            Width           =   180
         End
         Begin VB.Label lblEdit 
            Caption         =   $"frm��������.frx":0038
            Height          =   600
            Left            =   960
            TabIndex        =   55
            Top             =   2745
            Width           =   5865
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Index           =   9
            Left            =   372
            TabIndex        =   47
            Top             =   4350
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   8
            Left            =   372
            TabIndex        =   45
            Top             =   3930
            Width           =   540
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   5
            Left            =   3840
            TabIndex        =   43
            Top             =   2385
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ��"
            Height          =   180
            Index           =   7
            Left            =   372
            TabIndex        =   41
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   39
            Top             =   2010
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��ⵥ��"
            Height          =   180
            Index           =   6
            Left            =   192
            TabIndex        =   37
            Top             =   2016
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   3840
            TabIndex        =   35
            Top             =   1605
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   5
            Left            =   552
            TabIndex        =   33
            Top             =   1596
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   4
            Left            =   552
            TabIndex        =   31
            Top             =   1200
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���"
            Height          =   180
            Index           =   3
            Left            =   552
            TabIndex        =   29
            Top             =   840
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Ʒ��"
            Height          =   180
            Index           =   2
            Left            =   552
            TabIndex        =   27
            Top             =   456
            Width           =   360
         End
      End
      Begin VB.Frame fra��Χ 
         Height          =   5370
         Left            =   240
         TabIndex        =   53
         Top             =   432
         Width           =   6900
         Begin VB.CheckBox chkStorage 
            Caption         =   "������ҩƷ�������С�ڷ�Ʊ����(&S)"
            Height          =   270
            Left            =   564
            TabIndex        =   13
            Top             =   2100
            Width           =   4104
         End
         Begin VB.ComboBox cboStock 
            Height          =   300
            Left            =   1035
            TabIndex        =   21
            Top             =   4830
            Width           =   2460
         End
         Begin VB.TextBox txt����� 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   810
            Left            =   870
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   3885
            Width           =   5745
         End
         Begin VB.CheckBox chk����� 
            Caption         =   "��������Ų���(&F)"
            Height          =   384
            Left            =   570
            TabIndex        =   17
            Top             =   3555
            Width           =   1845
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   312
            Index           =   0
            Left            =   1632
            TabIndex        =   10
            Top             =   1704
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   314441731
            CurrentDate     =   36263
         End
         Begin VB.CheckBox chk��Ʊ���� 
            Caption         =   "��Ӧ�����ݵķ�Ʊ���ڲ���(&R)"
            Height          =   270
            Left            =   564
            TabIndex        =   8
            Top             =   1356
            Width           =   4104
         End
         Begin VB.TextBox txt��Ʊ�� 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   810
            Left            =   888
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   2685
            Width           =   5745
         End
         Begin VB.CommandButton Cmd��Ӧ�� 
            Caption         =   "��"
            Height          =   264
            Left            =   6435
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   348
            Width           =   255
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "��Ӧ�����ݵ�������ڲ���(&V)"
            Height          =   270
            Left            =   576
            TabIndex        =   3
            Top             =   720
            Width           =   4104
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   312
            Index           =   1
            Left            =   1632
            TabIndex        =   5
            Top             =   996
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   314572803
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   312
            Index           =   1
            Left            =   3540
            TabIndex        =   7
            Top             =   996
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   314572803
            CurrentDate     =   36263
         End
         Begin VB.CheckBox chk��Ʊ�� 
            Caption         =   "����Ʊ�Ų���(&F)"
            Height          =   384
            Left            =   564
            TabIndex        =   14
            Top             =   2370
            Width           =   1668
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   312
            Index           =   0
            Left            =   3540
            TabIndex        =   12
            Top             =   1704
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   314572803
            CurrentDate     =   36263
         End
         Begin VB.TextBox txt��Ӧ�� 
            Height          =   300
            Left            =   924
            MaxLength       =   50
            TabIndex        =   1
            Top             =   324
            Width           =   5520
         End
         Begin VB.Label lblStock 
            AutoSize        =   -1  'True
            Caption         =   "���ⷿ����"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   4905
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   ":������ڶ���������ݺţ����ö��ָ���"
            Height          =   180
            Index           =   10
            Left            =   2400
            TabIndex        =   18
            Top             =   3300
            Width           =   3330
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ����"
            Height          =   180
            Index           =   0
            Left            =   864
            TabIndex        =   9
            Top             =   1764
            Width           =   720
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   4
            Left            =   3300
            TabIndex        =   11
            Top             =   1764
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   ":������ڶ��ŷ�Ʊ�ţ����ö��ָ���"
            Height          =   180
            Index           =   0
            Left            =   2205
            TabIndex        =   15
            Top             =   2450
            Width           =   2970
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��Ӧ��(&G)"
            Height          =   180
            Index           =   1
            Left            =   72
            TabIndex        =   0
            Top             =   384
            Width           =   828
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   3300
            TabIndex        =   6
            Top             =   1056
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�������"
            Height          =   180
            Index           =   1
            Left            =   852
            TabIndex        =   4
            Top             =   1056
            Width           =   720
         End
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3090
      Left            =   1560
      TabIndex        =   51
      Top             =   4965
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   5450
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnAdvance As Boolean '�Ƿ�չ��
Private mlng��Ӧ��ID As Long
Private mstrSelectTag As String     '��ǰѡ��Ķ���
Private mstrFind As String
Private mstrPrivs As String
Private mblnOK As Boolean
Private mcllFilter As Collection
Private mrsStock As ADODB.Recordset
Private mlngModule As Long
Private mblnNoClick As Boolean

Public Sub Ȩ������()
    If Check���Ȩ��(mstrPrivs, "ҩƷ") = False Then
        chkType(0).Enabled = False
        chkType(0).Value = 0
    End If
    If Check���Ȩ��(mstrPrivs, "����") = False Then
        chkType(1).Enabled = False
        chkType(1).Value = 0
    End If
    
    If Check���Ȩ��(mstrPrivs, "�豸") = False Then
        chkType(2).Enabled = False
        chkType(2).Value = 0
    End If
    If Check���Ȩ��(mstrPrivs, "����") = False Then
        chkType(3).Enabled = False
        chkType(3).Value = 0
    End If
    If Check���Ȩ��(mstrPrivs, "����") = False Then
        chkType(4).Enabled = False
        chkType(4).Value = 0
    End If
End Sub

Public Function ShowFind(ByVal FrmMain As Form, ByVal lng��Ӧ��ID As Long, ByVal strPrivs As String, ByRef cllFilter As Collection, Optional int��� As Integer = 0) As Boolean
    '--------------------------------------------------------------
    '���ܣ���ȡ���������SQL���
    '������FrmMain-���ô���
    '       lng��Ӧ��ID-��Ӧ��ID
    '       strPrivs-Ȩ�޴�
    '���أ���������������true,���򷵻�False
    '˵����
    '--------------------------------------------------------------
    mstrFind = ""
    mstrPrivs = strPrivs
    '����27930 by lesfeng 2010-03-23
    If int��� = 1 Then Me.Caption = "��Ǹ�����������"
    If CheckCompete = False Then Exit Function
    mlng��Ӧ��ID = lng��Ӧ��ID
    Me.Show vbModal, FrmMain
    Set cllFilter = mcllFilter
    ShowFind = mblnOK
End Function
 

Private Sub cboStock_Click()
   If mblnNoClick Then Exit Sub
    If cboStock.ListIndex >= 0 Then cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    '����:33640
    If KeyAscii <> 13 Then Exit Sub
    If cboStock.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If mrsStock Is Nothing Then InitStockData
    If zlSelectDept(Me, mlngModule, cboStock, mrsStock, cboStock.Text, True, "���в���") = False Then KeyAscii = 0: Exit Sub
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_LostFocus()
    Dim i As Long
    If cboStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then
        For i = 0 To cboStock.ListCount - 1
            If Val(cboStock.Tag) = cboStock.ItemData(i) Then
                mblnNoClick = True
                cboStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub chkStorage_Click()
    Dim i As Integer
    If chkStorage.Value = 1 Then
        chkType(0).Value = 1
        For i = 1 To chkType.Count - 1
            chkType(i).Value = 0
        Next
    End If
End Sub

Private Sub chkType_Click(Index As Integer)
    If Index = 0 Then
        If chkType(Index) = 0 Then
            chkStorage.Value = 0
        End If
    Else
        If chkType(Index) = 1 Then
            chkStorage.Value = IIf(chkType(Index).Value = 1, 0, 1) And chkType(0).Value = 1
        End If
    End If
End Sub

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 4 Then
            If cmdȷ��.Enabled Then cmdȷ��.SetFocus
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Sub chk��Ʊ��_Click()
    txt��Ʊ��.Enabled = chk��Ʊ��.Value = 1
    txt��Ʊ��.BackColor = IIf(txt��Ʊ��.Enabled, vbWhite, Me.BackColor)
End Sub

Private Sub chk��Ʊ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chk�����_Click()
    txt�����.Enabled = chk�����.Value = 1
    txt�����.BackColor = IIf(txt�����.Enabled, vbWhite, Me.BackColor)
End Sub

Private Sub chk�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chk��Ʊ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chk��Ʊ����_Click()
    dtp��ʼʱ��(0).Enabled = IIf(chk��Ʊ����.Value = 1, True, False)
    dtp����ʱ��(0).Enabled = IIf(chk��Ʊ����.Value = 1, True, False)
End Sub

Private Sub chk���_Click()
    dtp��ʼʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
    dtp����ʱ��(1).Enabled = IIf(chk���.Value = 1, True, False)
End Sub
 
Private Sub Cmd��Ӧ��_Click()
    Dim strTemp As String
    
    strTemp = frm��Ӧ��ѡ��.SelDept(mstrPrivs)
    If strTemp = "" Then Exit Sub
    txt��Ӧ��.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    txt��Ӧ��.Tag = Left(strTemp, InStr(strTemp, ",") - 1)
    Unload frm��Ӧ��ѡ��
    If chk���.Enabled And chk���.Visible Then chk���.SetFocus
End Sub

Private Sub Cmdȡ��_Click()
    mblnOK = False
    Unload Me
End Sub

Private Function CheckValied() As Boolean
    '------------------------------------------------------------------------------------------------------
    '����:����������ݵĺϷ���
    '���:
    '����:
    '����: �Ϸ�����true,���򷵻�False
    '�޸���:���˺�
    '�޸�ʱ��:2007/2/28
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
 
    If chk��Ʊ��.Value = 1 And Trim(Replace(txt��Ʊ��.Text, vbCrLf, "")) = "" Then
        ShowMsgbox "δ���뷢Ʊ��,����!"
        sstFilter.Tab = 0
        If chk��Ʊ��.Enabled Then Me.chk��Ʊ��.SetFocus
        Exit Function
    End If
    
    If chk�����.Value = 1 And Trim(Replace(txt�����.Text, vbCrLf, "")) = "" Then
        ShowMsgbox "δ�����������,����!"
        sstFilter.Tab = 0
        If chk�����.Enabled Then Me.chk�����.SetFocus
        Exit Function
    End If
    
    If Check��Ӧ�� = False Then Exit Function
    
    For i = 0 To txtEdit.UBound
        If InStr(1, txtEdit(i).Text, "'") > 0 Then
            ShowMsgbox txtEdit(i).Tag & "���÷Ƿ��ַ�"
            sstFilter.Tab = 1
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
        If InStr(1, txtEdit(i).Tag, "��ʼ") > 0 Then
            If txtEdit(i).Text > txtEdit(i + 1).Text And txtEdit(i + 1).Text <> "" Then
                ShowMsgbox txtEdit(i).Tag & "����" & txtEdit(i + 1).Tag
                sstFilter.Tab = 1
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        End If
    Next
    Dim blnHaving As Boolean
    blnHaving = False
    For i = 0 To chkType.UBound
        If chkType(i).Value = 1 Then
            blnHaving = True
            Exit For
        End If
    Next
    If blnHaving = False Then
        ShowMsgbox "δѡ����Ҫ���ҵ����,����"
        sstFilter.Tab = 0
        If chkType(0).Enabled Then Me.chkType(0).SetFocus
        Exit Function
    End If
    
    CheckValied = True
End Function

Private Sub Cmdȷ��_Click()
    Dim strTemp As String
    Dim i As Long
    mstrFind = ""
    '����SQL��ѯ�������
    Dim intTemp As Integer
    Dim strFind As String
    
    If CheckValied = False Then Exit Sub
    
    Set mcllFilter = New Collection
    mcllFilter.Add Val(txt��Ӧ��.Tag), "��Ӧ��ID"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "�������"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "��Ʊ����"
    mcllFilter.Add "", "��Ʊ���б�"
    mcllFilter.Add "", "��������б�"
    mcllFilter.Add "", "ϵͳ��ʶ"
    mcllFilter.Add "", "Ʒ��"
    mcllFilter.Add "", "���"
    mcllFilter.Add "", "����"
    mcllFilter.Add Array("", ""), "����"
    mcllFilter.Add Array("", ""), "��ⵥ��"
    mcllFilter.Add Array("", ""), "��Ʊ��"
    mcllFilter.Add Array("", ""), "�������"
    mcllFilter.Add "", "������"
    mcllFilter.Add "", "�����"
    mcllFilter.Add 0, "�ⷿ"
    mcllFilter.Add "0", "������ҩƷ�������С�ڷ�Ʊ����"
    
    If chk��Ʊ����.Value = 1 And chk���.Value = 1 Then
        mstrFind = " And ( ([alias]������� Between [2] And [3]) or ([alias]��Ʊ���� Between [4] And [5])) "
        mcllFilter.Remove "�������"
        mcllFilter.Remove "��Ʊ����"
        mcllFilter.Add Array(Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & " 00:00:00", Format(dtp����ʱ��(1), "yyyy-mm-dd") & " 23:59:59"), "�������"
        mcllFilter.Add Array(Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & " 00:00:00", Format(dtp����ʱ��(0), "yyyy-mm-dd") & " 23:59:59"), "��Ʊ����"
    ElseIf chk���.Value = 1 Then
        mstrFind = " And ( [alias]������� Between [2] And [3]) "
        mcllFilter.Remove "�������"
        mcllFilter.Add Array(Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & " 00:00:00", Format(dtp����ʱ��(1), "yyyy-mm-dd") & " 23:59:59"), "�������"
    ElseIf chk��Ʊ����.Value = 1 Then
        mstrFind = " And ([alias]��Ʊ���� Between [4] And [5]) "
        mcllFilter.Remove "��Ʊ����"
        mcllFilter.Add Array(Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & " 00:00:00", Format(dtp����ʱ��(0), "yyyy-mm-dd") & " 23:59:59"), "��Ʊ����"
    End If
    
    '������ҩƷ�������С�ڷ�Ʊ����
    mcllFilter.Remove "������ҩƷ�������С�ڷ�Ʊ����"
    mcllFilter.Add chkStorage.Value, "������ҩƷ�������С�ڷ�Ʊ����"
    
    If chk��Ʊ��.Value = 1 Then
        mstrFind = mstrFind & " And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                              "             From Table(Cast(f_Str2list([6]) As zlTools.t_Strlist)) J " & vbCrLf & _
                              "             Where Instr(',' || [alias]��Ʊ�� || ',', ',' || Column_Value || ',') > 0) "
        
        mcllFilter.Remove "��Ʊ���б�"
        mcllFilter.Add Replace(txt��Ʊ��.Text, vbCrLf, ""), "��Ʊ���б�"
    End If
    If chk�����.Value = 1 Then
        mstrFind = mstrFind & " And exists(Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                              "            From Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) J " & vbCrLf & _
                              "            Where J.Column_Value = [alias]�������) "
        mcllFilter.Remove "��������б�"
        mcllFilter.Add Replace(txt�����.Text, vbCrLf, ""), "��������б�"
    End If
    'ȷ��������
    Dim blnAll As Boolean
    blnAll = True
    strTemp = ""
    For i = 0 To chkType.UBound
        If chkType(i).Value = 1 Then
            strTemp = strTemp & "," & i + 1
        Else
          blnAll = False
        End If
    Next
    strTemp = Mid(strTemp, 2)
    If blnAll = False Then
        '1����ҩƷӦ����   2��������Ӧ����   3�����豸Ӧ����   4��������,5--��������
        mstrFind = mstrFind & _
                " And exists(Select /*+ cardinality(J, 10)*/ 1 " & _
                "            From Table(Cast(f_Num2list([8]) As zlTools.t_Numlist)) J " & _
                "            Where J.Column_Value = [alias]ϵͳ��ʶ )"
        mcllFilter.Remove "ϵͳ��ʶ"
        mcllFilter.Add strTemp, "ϵͳ��ʶ"
    End If
    If cboStock.ListIndex >= 0 Then
        If cboStock.ItemData(cboStock.ListIndex) <> 0 Then
            mstrFind = mstrFind & " And [alias]�ⷿID=[23]"
            mcllFilter.Remove "�ⷿ"
            mcllFilter.Add cboStock.ItemData(cboStock.ListIndex), "�ⷿ"
        End If
    End If
    
    '��չ��ѯ����
    If mblnAdvance = False Then
        GoTo EndSub:
    End If
    '------------------------------------------------------------------------------------------------------------
    'Ʒ��
    If Trim(txtEdit(0).Text) <> "" Then
        strTemp = GetMatchingSting(Trim(txtEdit(0).Text), False)
        mcllFilter.Remove "Ʒ��"
        mcllFilter.Add strTemp, "Ʒ��"
        
        strFind = " And [alias]Ʒ�� like [9]"
        If zlCommFun.IsCharAlpha(Trim(txtEdit(0).Text)) Then          '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            '.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ", , , True))
            If gSystemPara.int���뷽ʽ = 1 Then
                '������ѯ
                If Mid(gSystemPara.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And zltools.zlWBCode([alias]Ʒ��) Like upper([9]) "
            ElseIf gSystemPara.int���뷽ʽ = 0 Then
                If Mid(gSystemPara.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And zltools.zlspellcode([alias]Ʒ��) Like Upper([9]) "
            Else
                If Mid(gSystemPara.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And (zltools.zlWBCode([alias]Ʒ��) Like Upper([9]) or zltools.zlspellcode([alias]Ʒ��) Like upper([9]) "
            End If
        End If
        mstrFind = mstrFind & strFind
    End If
    '------------------------------------------------------------------------------------------------------------
    '���
    If Trim(txtEdit(1).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]��� like [10]"
        strTemp = GetMatchingSting(Trim(txtEdit(1).Text), False)
        mcllFilter.Remove "���"
        mcllFilter.Add strTemp, "���"
    End If
    '����
    If Trim(txtEdit(2).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]��� like [11]"
        strTemp = GetMatchingSting(Trim(txtEdit(2).Text), False)
        mcllFilter.Remove "����"
        mcllFilter.Add strTemp, "����"
    End If
    '����
    If Trim(txtEdit(3).Text) <> "" And Trim(txtEdit(4).Text) <> "" Then
        mstrFind = mstrFind & " And ([alias]���� between [12] and [13])"
    ElseIf Trim(txtEdit(3).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]���� >= [12] "
    ElseIf Trim(txtEdit(4).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]���� <= [13] "
    End If
    mcllFilter.Remove "����"
    mcllFilter.Add Array(Trim(txtEdit(3).Text), Trim(txtEdit(4).Text)), "����"
    '��ⵥ��
    If Trim(txtEdit(5).Text) <> "" And Trim(txtEdit(6).Text) <> "" Then
        mstrFind = mstrFind & " And ([alias]��ⵥ�ݺ� between [14] and [15])"
    ElseIf Trim(txtEdit(5).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]��ⵥ�ݺ� >= [14] "
    ElseIf Trim(txtEdit(6).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]��ⵥ�ݺ� <= [15] "
    End If
    mcllFilter.Remove "��ⵥ��"
    mcllFilter.Add Array(Trim(txtEdit(5).Text), Trim(txtEdit(6).Text)), "��ⵥ��"
    '��Ʊ��
    If Trim(txtEdit(7).Text) <> "" And Trim(txtEdit(8).Text) <> "" Then
        mstrFind = mstrFind & " And ([alias]��Ʊ�� between [16] and [17])"
    ElseIf Trim(txtEdit(7).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]��Ʊ�� >= [16] "
    ElseIf Trim(txtEdit(8).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]��Ʊ�� <= [17] "
    End If
    mcllFilter.Remove "��Ʊ��"
    mcllFilter.Add Array(Trim(txtEdit(7).Text), Trim(txtEdit(8).Text)), "��Ʊ��"
    '�������
    If Trim(txtEdit(11).Text) <> "" And Trim(txtEdit(12).Text) <> "" Then
        mstrFind = mstrFind & " And ([alias]������� between [18] and [19])"
    ElseIf Trim(txtEdit(11).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]������� >= [18] "
    ElseIf Trim(txtEdit(12).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]������� <= [19] "
    End If
    mcllFilter.Remove "�������"
    mcllFilter.Add Array(Trim(txtEdit(11).Text), Trim(txtEdit(12).Text)), "�������"
    '������:
    If Trim(txtEdit(9).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]������ like [20] "
        mcllFilter.Remove "������"
        mcllFilter.Add Trim(txtEdit(9).Text), "������"
    End If
    '�����:
    If Trim(txtEdit(10).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]����� like [21] "
        mcllFilter.Remove "�����"
        mcllFilter.Add Trim(txtEdit(10).Text), "�����"
    End If
EndSub:
    mcllFilter.Add mstrFind, "����"
    
    mblnOK = True
    Unload Me
End Sub

Private Sub dtp����ʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
     End If
End Sub

Private Sub dtp��ʼʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
     End If
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mlngModule = 1323
    '����:Ȩ�޿���:2008-08-18 14:41:40
    Call Ȩ������
    '����27930 by lesfeng 2010-03-23
    Call setInitDate
    
    Call InitStockData  '33640
    
    Me.txt��Ӧ��.Tag = 0
    On Error GoTo errHandle
    If mlng��Ӧ��ID <> 0 Then
        gstrSQL = "Select ����,���� from ��Ӧ�� where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng��Ӧ��ID)
        If rsTemp.EOF = False Then
            txt��Ӧ��.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
            txt��Ӧ��.Tag = mlng��Ӧ��ID
        End If
    End If
    '�򿪼�¼��
    sstFilter.Tab = 0
    mblnAdvance = False
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function CheckCompete() As Boolean
    '--------------------------------------------------------------
    '���ܣ�����Ƿ��й�Ӧ������
    '������
    '���أ��Ƿ��й�Ӧ������
    '˵����
    '--------------------------------------------------------------
    Dim rsCompete As New Recordset
    
    CheckCompete = False
    Err = 0
    On Error GoTo ErrHand:
    gstrSQL = "Select id From ��Ӧ�� Where (����ʱ�� is null or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01') " & zl_��ȡվ������ & "  and  ĩ��=1 and rownum<=2 "
    zlDatabase.OpenRecordset rsCompete, gstrSQL, "��鹩Ӧ��"
    With rsCompete
        If .EOF Then
            .Close
            ShowMsgbox "��Ӧ����Ϣ��ȫ�����ڹ�Ӧ�̹��������ù�Ӧ����Ϣ��"
            Exit Function
        End If
    End With
    CheckCompete = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            Case "Provider"
                txt��Ӧ��.SetFocus
            Case "Booker"
                txtEdit(9).SetFocus
            Case "Verify"
                txtEdit(10).SetFocus
        End Select
        Cancel = True
    End If
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Provider"
                    txt��Ӧ��.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
                    txt��Ӧ��.Tag = .TextMatrix(.Row, 0)
                    If chk���.Enabled And chk���.Visible Then chk���.SetFocus
                Case "9"
                    txtEdit(9) = .TextMatrix(.Row, 2)
                    If txtEdit(10).Enabled And txtEdit(10).Visible Then txtEdit(10).SetFocus
                Case "10"
                    txtEdit(10) = .TextMatrix(.Row, 2)
                    If cmdȷ��.Enabled Then cmdȷ��.SetFocus
            End Select
            .Visible = False
            Exit Sub
        End If
    End With
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    With sstFilter
        If .Tab = 1 Then
            mblnAdvance = True
        End If
    End With
End Sub

Private Sub sstFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        If sstFilter.Tab = 0 Then
          If txt��Ӧ��.Enabled Then txt��Ӧ��.SetFocus
        Else
            If txtEdit(0).Enabled Then txtEdit(0).SetFocus
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If InStr(1, txtEdit(Index).Tag, "��") > 0 Then
         SetTxtGotFocus txtEdit(Index), True
    Else
        zlControl.TxtSelAll txtEdit(Index)
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If InStr(1, txtEdit(Index).Tag, "��") <> 0 Then
            Call SelectPerson(Index, KeyCode)
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Function SelectPerson(ByVal intIndex As Integer, ByRef KeyCode As Integer) As Boolean
    '����:ѡ����ص���Ա��Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    If Trim(txtEdit(intIndex).Text) = "" Then
        zlCommFun.PressKey vbKeyTab
        Exit Function
    End If
    txtEdit(intIndex).Text = UCase(txtEdit(intIndex).Text)
    strKey = txtEdit(intIndex).Text
    strKey = GetMatchingSting(strKey)
    
    On Error GoTo errHandle
    gstrSQL = "" & _
        "   Select ���,����,���� " & _
        "   From ��Ա�� " & _
        "   Where (���� like [1] or ��� like [1] or ���� like [1] )  " & zl_��ȡվ������ & "" & _
        "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
        "   order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ" & txtEdit(intIndex).Tag, strKey)
       
    With rsTemp
        If .EOF Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            KeyCode = 0
            txtEdit(intIndex).SelStart = 0
            txtEdit(intIndex).SelLength = Len(txtEdit(intIndex).Text)
            
            Exit Function
        End If
        Dim sngHeight As Single
        
        If .RecordCount > 1 Then
            mstrSelectTag = intIndex
            
            Set mshSelect.Recordset = rsTemp
            sngHeight = sstFilter.Top + fra.Top + txtEdit(intIndex).Top
            If sngHeight > mshSelect.Rows * (mshSelect.RowHeight(0) + 30) + 200 Then
                mshSelect.Height = mshSelect.Rows * (mshSelect.RowHeight(0) + 30) + 200
            Else
                mshSelect.Height = sngHeight
            End If
            With mshSelect
                .Top = sstFilter.Top + fra.Top + txtEdit(intIndex).Top - .Height
                .Left = sstFilter.Left + fra.Left + txtEdit(intIndex).Left
                .Visible = True
                .SetFocus
                .ColWidth(0) = 800
                .ColWidth(1) = 1500
                .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                .ZOrder 0
                Exit Function
            End With
        Else
            txtEdit(intIndex).Text = IIf(IsNull(!����), "", !����)
        End If
    End With
    zlCommFun.PressKey vbKeyTab
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    ImeLanguage False
End Sub

Private Sub txt��Ʊ��_GotFocus()
    zlControl.TxtSelAll txt��Ʊ��
    zlCommFun.OpenIme False
End Sub

Private Sub txt��Ʊ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt��Ʊ��_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��Ʊ��, KeyAscii, m�ı�ʽ
End Sub

Private Function Check��Ӧ��() As Boolean
    '---------------------------------------------------------------------------------------------
    '����:�����صĹ�Ӧ��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/11/05
    '---------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long
    
    Dim strVarr As Variant, strTemp As String, str��ʼ��Ʊ�� As String, str������Ʊ�� As String
    Err = 0: On Error GoTo ErrHand:
    
    If Val(txt��Ӧ��.Tag) <> 0 Then
        Check��Ӧ�� = True
        Exit Function
    End If
    
    If chk��Ʊ��.Value = 1 Then
        If chk���.Value = 0 Then
            ShowMsgbox "����δѡ����ص����ʱ���Ӧ�̣�Ϊ��������ܣ������ѡ��һ��������������ڻ�Ӧ�̣���!"
            Exit Function
        End If
        
        strVarr = Split(Replace(txt��Ʊ��.Text, vbCrLf, ""), ",")
        strTemp = ""
        For i = 0 To UBound(strVarr)
            If Trim(strVarr(i)) <> "" Then
                strTemp = strTemp & "," & strVarr(i)
            End If
        Next
        strTemp = Mid(strTemp, 2)
        If strTemp = "" Then
            ShowMsgbox "δ������صķ�Ʊ��,����!"
            Exit Function
        End If
        str��ʼ��Ʊ�� = strTemp
         
        If InStr(1, strTemp, ",") <> 0 Then
            gstrSQL = "" & _
                "   Select distinct M.id,M.����,M.����,M.ĩ��,M.����,M.���֤��,M.���֤Ч��,M.ִ�պ�,M.ִ��Ч��," & _
                "           M.˰��ǼǺ�,M.��ַ,M.��������,M.�ʺ�,M.��ϵ��,M.����ʱ��,M.����,M.������ " & _
                "   From Ӧ����¼ A,��Ӧ�� M" & _
                "   Where A.��λID=M.ID and a.������� between [3] and [4] " & vbCrLf & _
                "         And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                "                     From Table(Cast(f_Str2list(A.��Ʊ��) As zlTools.t_Strlist)) J " & vbCrLf & _
                "                     Where exists(Select /*+ cardinality(M, 10)*/ 1 " & vbCrLf & _
                "                                  From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) M " & vbCrLf & _
                "                                  Where j.Column_Value=m.Column_Value)) "
        Else
            gstrSQL = "" & _
                "   Select  distinct M.id,M.����,M.����,M.ĩ��,M.����,M.���֤��,M.���֤Ч��,M.ִ�պ�,M.ִ��Ч��," & _
                "           M.˰��ǼǺ�,M.��ַ,M.��������,M.�ʺ�,M.��ϵ��,M.����ʱ��,M.����,M.������ " & _
                "   From Ӧ����¼ A,��Ӧ�� M  " & _
                "   Where  A.��λID=M.ID and a.������� between [3] and [4] " & _
                "         And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                "                     From Table(Cast(f_Str2list(A.��Ʊ��) As zlTools.t_Strlist)) J " & vbCrLf & _
                "                     Where j.Column_Value=[1])"
                
        End If
    ElseIf Trim(txtEdit(7).Text) <> "" Or txtEdit(8).Text <> "" Then
        If chk���.Value = 0 Then
            ShowMsgbox "����δѡ����ص����ʱ���Ӧ�̣�Ϊ��������ܣ������ѡ��һ��������������ڻ�Ӧ�̣���!"
            Exit Function
        End If
        strTemp = ""
        str��ʼ��Ʊ�� = Trim(txtEdit(7).Text)
        str������Ʊ�� = Trim(txtEdit(8).Text)
        
        If Trim(txtEdit(7).Text) <> "" And Trim(txtEdit(8).Text) <> "" Then
            strTemp = "  And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                      "              From Table(Cast(f_Str2list(A.��Ʊ��) As zlTools.t_Strlist)) J " & vbCrLf & _
                      "              Where j.Column_Value>=[1]  and j.Column_Value<=[2])"
        ElseIf Trim(txtEdit(7).Text) = "" And Trim(txtEdit(8).Text) <> "" Then
'            strTemp = "  And A.��Ʊ��<=[2] "
            strTemp = "  And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                      "              From Table(Cast(f_Str2list(A.��Ʊ��) As zlTools.t_Strlist)) J " & vbCrLf & _
                      "              Where  j.Column_Value<=[2])"
        Else
            strTemp = "  And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                      "              From Table(Cast(f_Str2list(A.��Ʊ��) As zlTools.t_Strlist)) J " & vbCrLf & _
                      "              Where  j.Column_Value>=[1])"
        End If
        
        gstrSQL = "" & _
            "   Select  distinct M.id,M.����,M.����,M.ĩ��,M.����,M.���֤��,M.���֤Ч��,M.ִ�պ�,M.ִ��Ч��," & _
            "           M.˰��ǼǺ�,M.��ַ,M.��������,M.�ʺ�,M.��ϵ��,M.����ʱ��,M.����,M.������ " & _
            "   From  Ӧ����¼ A,��Ӧ�� M" & _
            "   Where  A.��λID=M.ID And a.������� between [3] and [4]  " & strTemp
    Else
        If Val(txt��Ӧ��.Tag) = 0 Then
            ShowMsgbox "��Ӧ��δѡ��,���ܼ���!"
            sstFilter.Tab = 0
            If txt��Ӧ��.Enabled Then txt��Ӧ��.SetFocus
        Else
            Check��Ӧ�� = True
        End If
        Exit Function
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str��ʼ��Ʊ��, str������Ʊ��, _
           CDate(Format(dtp��ʼʱ��(1).Value, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(dtp����ʱ��(1).Value, "yyyy-mm-dd") & " 23:59:59"))
    If rsTemp.EOF = True Then
        ShowMsgbox "�޴˷�Ʊ�ŵĹ�Ӧ��,����!"
        Exit Function
    End If
    txt��Ӧ��.Text = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
    txt��Ӧ��.Tag = Nvl(rsTemp!ID)
    Check��Ӧ�� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txt��Ӧ��_Change()
    txt��Ӧ��.Tag = ""
End Sub

Private Sub txt��Ӧ��_GotFocus()
    SetTxtGotFocus txt��Ӧ��, True
End Sub

Private Function Select��Ӧ��(ByVal strKey As String) As Boolean
    '----------------------------------------------------------------------------------------
    '����:ѡ��Ӧ��
    '����:strKey-ѡ��Ӧ��
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/11/5
    '----------------------------------------------------------------------------------------
    Dim strȨ�� As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Err = 0: On Error GoTo ErrHand:
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
    End If
      
    strȨ�� = " and " & Get����Ȩ��(mstrPrivs)
    gstrSQL = "" & _
        "   Select id, ����,����,ĩ��,����,���֤��,���֤Ч��,ִ�պ�,ִ��Ч��,˰��ǼǺ�,��ַ,��������,�ʺ�,��ϵ��,����ʱ��,����,������" & _
        "   From ��Ӧ�� " & _
        "   where   ĩ��=1 " & zl_��ȡվ������ & "  " & _
        "           and  (����ʱ�� is null or ����ʱ��>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & _
        "           and (���� like [1] or ���� like [1] or ���� like [1])  " & strȨ��
    'ShowSelect:
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    Dim lngX As Long, lngY As Long, lngH As Long
    lngX = Me.Left + txt��Ӧ��.Left + Screen.TwipsPerPixelX
    lngY = Me.Top + Me.Height - Me.ScaleHeight + txt��Ӧ��.Top
    lngH = txt��Ӧ��.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "��Ӧ��ѡ��", False, "", "ѡ��Ӧ��", False, True, True, lngX, lngY, lngH, blnCancel, False, True, strKey)
    If blnCancel Then Exit Function
    If rsTemp Is Nothing Then
        ShowMsgbox "�����ڷ��������Ĺ�Ӧ��,����!"
        Exit Function
    End If
    If rsTemp.State <> 1 Then Exit Function
    txt��Ӧ�� = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
    txt��Ӧ��.Tag = Nvl(rsTemp!ID)
    Select��Ӧ�� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New Recordset
    Dim strTemp As String
    Dim strȨ�� As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub '
    If Val(txt��Ӧ��.Tag) = 0 Then
        If Trim(txt��Ӧ��.Text) <> "" Then
            If Select��Ӧ��(Trim(txt��Ӧ��.Text)) = False Then
                Exit Sub
            End If
        End If
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
   
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt�����_GotFocus()
    zlControl.TxtSelAll txt�����
    zlCommFun.OpenIme False
End Sub

Private Sub txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt�����, KeyAscii, m�ı�ʽ)
End Sub
'����27930 by lesfeng 2010-03-23
Private Sub setInitDate()
    Dim arrHead As Variant
    Dim strMonth As String
    Dim intGetEndMonth As Integer
    Dim intGetBeginMonth As Integer
    Dim blnMonth As Boolean
    Dim dtTempDate As Date
    Dim dtTemp As Date
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strMonth = zlDatabase.GetPara("���ø���ʱ��", glngSys, 1323)
    If InStr(1, strMonth, "-") > 0 Then
        arrHead = Split(strMonth, "-")
        blnMonth = Val(arrHead(0)) = 1
        intGetEndMonth = Val(arrHead(1))
        intGetBeginMonth = Val(arrHead(2))
    Else
        blnMonth = False
        intGetEndMonth = 0
        intGetBeginMonth = 0
    End If
    On Error GoTo errHandle
    If blnMonth Then
        dtTempDate = DateAdd("m", -intGetEndMonth, zlDatabase.Currentdate)
        dtTemp = CDate(Format(dtTempDate, "yyyy-MM") & "-01")
        strSQL = "select to_date('" & Format(dtTemp, "yyyy-mm-dd") & "','yyyy-mm-dd') -1/24/60/60 as dtdate from dual"
        zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
        If Not rsTemp.EOF Then
            Me.dtp����ʱ��(0) = IIf(IsNull(rsTemp!dtdate), zlDatabase.Currentdate, rsTemp!dtdate)
            If intGetBeginMonth = 0 Then
                Me.dtp��ʼʱ��(0) = Me.dtp����ʱ��(0)
            Else
                Me.dtp��ʼʱ��(0) = DateAdd("m", -intGetBeginMonth, dtTemp)
            End If
        Else
            Me.dtp����ʱ��(0) = zlDatabase.Currentdate
            Me.dtp��ʼʱ��(0) = DateAdd("d", -7, Me.dtp����ʱ��(0))
        End If
        rsTemp.Close
    Else
        Me.dtp����ʱ��(0) = zlDatabase.Currentdate
        Me.dtp��ʼʱ��(0) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    End If
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    Me.dtp��ʼʱ��(1) = Me.dtp��ʼʱ��(0)
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub InitStockData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؿⷿ����
    '����:���˺�
    '����:2010-11-02 16:13:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strStock As String
    strStock = "HIJKLMN"
    strStock = strStock & "V"   '���Ŀ���Ƽ���(K)
    strStock = strStock & "RS"  '���ʿⷿ�͹�Ӧ��
    strStock = strStock & "T"   '�豸��
    On Error GoTo errHandle
    gstrSQL = "" & _
    "   SELECT DISTINCT a.id,A.����, a.����,A.���� " & _
    "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
    "   Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� is Null) And c.�������� = b.���� " & _
    "           AND Instr([1],b.����,1) > 0 " & _
    "           AND a.id = c.����id " & _
    "           AND TO_CHAR(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
    "   Order by ����"
    Set mrsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock)
    With mrsStock
        cboStock.Clear
        cboStock.AddItem "���пⷿ"
        cboStock.ListIndex = cboStock.NewIndex
        Do While Not .EOF
            cboStock.AddItem Nvl(!����) & IIf(Nvl(!����) = "", "", "-") & Nvl(!����)
            cboStock.ItemData(cboStock.NewIndex) = Val(Nvl(!ID))
            If cboStock.ListIndex < 0 Then cboStock.ListIndex = cboStock.NewIndex
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


