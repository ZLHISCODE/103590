VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmAuditLogManage 
   BackColor       =   &H80000005&
   Caption         =   "��Ҫ�����䶯��־����"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   ControlBox      =   0   'False
   Icon            =   "frmAuditLogManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmAuditLogManage.frx":6852
   ScaleHeight     =   7305
   ScaleWidth      =   10500
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList img16 
      Left            =   9315
      Top             =   2580
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
            Picture         =   "frmAuditLogManage.frx":6D4B
            Key             =   "system"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditLogManage.frx":D5AD
            Key             =   "program"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgModuleType 
      Left            =   120
      Top             =   3015
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   51
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditLogManage.frx":13E0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditLogManage.frx":1735C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAuditLogConfig 
      BackColor       =   &H80000005&
      FillColor       =   &H80000000&
      Height          =   6060
      Left            =   1110
      ScaleHeight     =   6000
      ScaleWidth      =   8175
      TabIndex        =   4
      Top             =   5955
      Width           =   8235
      Begin MSComctlLib.TreeView tvwAuditLogConfig 
         Height          =   3030
         Left            =   90
         TabIndex        =   32
         Top             =   495
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   5345
         _Version        =   393217
         Indentation     =   706
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "ͣ��(&T)"
         Height          =   350
         Left            =   4410
         TabIndex        =   24
         Top             =   45
         Width           =   1100
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   3315
         TabIndex        =   23
         Top             =   45
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   585
         TabIndex        =   21
         Top             =   75
         Width           =   2000
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAuditLogConfig 
         Height          =   1680
         Left            =   2760
         TabIndex        =   18
         Top             =   495
         Width           =   4770
         _cx             =   8414
         _cy             =   2963
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   260
         RowHeightMax    =   260
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAuditLogManage.frx":1A852
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         ExplorerBar     =   1
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
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����"
         Height          =   180
         Left            =   105
         TabIndex        =   22
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox picAuditLogList 
      BackColor       =   &H80000005&
      Height          =   5955
      Left            =   915
      ScaleHeight     =   5895
      ScaleWidth      =   7755
      TabIndex        =   3
      Top             =   225
      Width           =   7815
      Begin VB.Frame fraDescription 
         BackColor       =   &H80000005&
         Caption         =   "����˵��"
         Height          =   2300
         Left            =   4425
         TabIndex        =   20
         Top             =   3630
         Width           =   3150
         Begin VB.TextBox txtInstructions 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   1920
            Left            =   135
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   255
            Width           =   2835
         End
      End
      Begin VB.Frame fraNote 
         BackColor       =   &H80000005&
         Caption         =   "��������"
         Height          =   2300
         Left            =   345
         TabIndex        =   19
         Top             =   2790
         Width           =   3150
         Begin VB.TextBox txtOperationContent 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   1920
            Left            =   135
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   255
            Width           =   2835
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAuditLogList 
         Height          =   2730
         Left            =   -10
         TabIndex        =   17
         Top             =   -10
         Width           =   3360
         _cx             =   5927
         _cy             =   4815
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   260
         RowHeightMax    =   260
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAuditLogManage.frx":1A950
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
         ExplorerBar     =   1
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
      Begin VB.PictureBox picFind 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3540
         Left            =   4125
         ScaleHeight     =   3540
         ScaleWidth      =   3495
         TabIndex        =   11
         Top             =   0
         Width           =   3500
         Begin VB.ListBox lisShowList 
            Appearance      =   0  'Flat
            Height          =   1470
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   420
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.ComboBox cboFunction 
            Height          =   300
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1659
            Width           =   2385
         End
         Begin VB.ComboBox cboSystem 
            Height          =   300
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   897
            Width           =   2385
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "����(&F)"
            Height          =   350
            Left            =   2175
            TabIndex        =   10
            Top             =   3135
            Width           =   1100
         End
         Begin VB.ComboBox cboWorkStation 
            Height          =   300
            Left            =   915
            TabIndex        =   5
            Top             =   135
            Width           =   2385
         End
         Begin VB.ComboBox cboModule 
            Height          =   300
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1278
            Width           =   2385
         End
         Begin VB.ComboBox cboUserName 
            Height          =   300
            Left            =   915
            TabIndex        =   6
            Top             =   516
            Width           =   2385
         End
         Begin MSComCtl2.DTPicker dtpDateEnd 
            Height          =   315
            Left            =   915
            TabIndex        =   9
            Top             =   2700
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   105381891
            CurrentDate     =   37029
         End
         Begin MSComCtl2.DTPicker dtpDateStart 
            Height          =   315
            Left            =   915
            TabIndex        =   8
            Top             =   2040
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   105381891
            CurrentDate     =   37029
         End
         Begin VB.Label lblFunction 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   150
            TabIndex        =   31
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label lblSystem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ϵͳ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   150
            TabIndex        =   29
            Top             =   945
            Width           =   720
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Left            =   915
            TabIndex        =   16
            Top             =   2430
            Width           =   180
         End
         Begin VB.Label lblWorkStation 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ͻ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   330
            TabIndex        =   15
            Top             =   210
            Width           =   540
         End
         Begin VB.Label lblModule 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ģ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   150
            TabIndex        =   14
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label lblUserName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�û���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   330
            TabIndex        =   13
            Top             =   570
            Width           =   540
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   165
            TabIndex        =   12
            Top             =   2085
            Width           =   720
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   1395
      Left            =   60
      TabIndex        =   2
      Top             =   615
      Width           =   1560
      _Version        =   589884
      _ExtentX        =   2752
      _ExtentY        =   2461
      _StockProps     =   64
   End
   Begin VB.CommandButton cmdLogClear 
      Caption         =   "��־����(&C)"
      Height          =   375
      Left            =   8430
      TabIndex        =   1
      Top             =   150
      Width           =   1290
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������־����"
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
      Left            =   195
      TabIndex        =   0
      Top             =   135
      Width           =   1440
   End
End
Attribute VB_Name = "frmAuditLogManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsAuditLog As ADODB.Recordset '��¼��������ѯ��������־���ݣ���Ҫ����չʾ��ϸ�������ݺͲ���˵��
Private mrsModuleList As ADODB.Recordset '��¼ģ�鼰����ͣ��Ϣ����Ҫ����ģ��Ĳ���
Private mrsWorkStation As ADODB.Recordset '��¼�ͻ�����Ϣ����Ҫ���ڿͻ���ģ������
Private mrsUserName As ADODB.Recordset '��¼�û�����Ϣ����Ҫ�����û�����ģ������
Private mrsSysProgFun As ADODB.Recordset '��¼ϵͳ��ģ�鼰���ܵĶ�Ӧ��ϵ
Private mlngCurPos As Long '��ǰ�������νṹ��λ��

Private Enum AuditLogList
    VLL_�û��� = 0
    VLL_��Ա = 1
    VLL_���� = 2
    VLL_����վ = 3
    VLL_�������� = 4
    VLL_ϵͳ��� = 5
    VLL_����ϵͳ = 6
    VLL_����ģ���� = 7
    VLL_����ģ�� = 8
    VLL_�������� = 9
    VLL_����ʱ�� = 10
    VLL_�������� = 11
    VLL_����˵�� = 12
    VLL_��ϸ�������� = 13
    VLL_��ϸ����˵�� = 14
End Enum

Private Enum AuditLogConfig
    VLC_ID = 0
    VLC_����ϵͳ = 1
    VLC_ģ������ = 2
    VLC_�������� = 3
    VLC_˵�� = 4
    VLC_����� = 5
    VLC_״̬ = 6
    VLC_״̬��� = 7
End Enum

Private Sub cboFunction_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '����س�������Ƶ���һ���ؼ�
        dtpDateStart.SetFocus
    End If
End Sub

Private Sub cboModule_Click()
    If cboModule.Text = "" Then
        '��ģ��ѡ��Ϊ�գ����ڲ���������������չʾ��Ӧϵͳ�е�ȫ����������
        If cboSystem.Text = "" Then
            mrsSysProgFun.Filter = ""
        Else
            mrsSysProgFun.Filter = "ϵͳ = " & Split(cboSystem.Text, "-")(0)
        End If
    Else
        '��ģ��ѡ��Ϊ�գ����ڲ��������������н�չʾ��ģ���е�ȫ����������
        If cboSystem.Text = "" Then
            mrsSysProgFun.Filter = "ģ�� = " & Split(cboModule.Text, "-")(0)
        Else
            mrsSysProgFun.Filter = "ϵͳ = " & Split(cboSystem.Text, "-")(0) & " And ģ�� = " & Split(cboModule.Text, "-")(0)
        End If
    End If
    
    '��������������
    mrsSysProgFun.Sort = "Id"
    cboFunction.Clear
    cboFunction.addItem ""
    Do While Not mrsSysProgFun.EOF
        cboFunction.addItem mrsSysProgFun!id & "-" & mrsSysProgFun!����
        mrsSysProgFun.MoveNext
    Loop
End Sub

Private Sub cboModule_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '����س�������Ƶ���һ���ؼ�
        cboFunction.SetFocus
    End If
End Sub

Private Sub cboSystem_Click()
    Dim strLastModuleNo As String

    If cboSystem.Text = "" Then
        '��ϵͳѡ��Ϊ�գ����ڲ���ģ�鼰����������������չʾȫ��������
        mrsSysProgFun.Filter = ""
    Else
        '��ϵͳѡ��ΪӦ��ϵͳ����ô�ڲ���ģ�鼰���������������н�չʾ��Ӧϵͳ�������
        mrsSysProgFun.Filter = "ϵͳ = " & Split(cboSystem.Text, "-")(0)
    End If
    
    '������ģ������
    mrsSysProgFun.Sort = "ϵͳ, ģ��"
    cboModule.Clear
    cboModule.addItem ""
    Do While Not mrsSysProgFun.EOF
        If mrsSysProgFun!ϵͳ & "-" & mrsSysProgFun!ģ�� <> strLastModuleNo Then
            cboModule.addItem mrsSysProgFun!ģ�� & "-" & mrsSysProgFun!ģ������
            strLastModuleNo = mrsSysProgFun!ϵͳ & "-" & mrsSysProgFun!ģ��
        End If
        mrsSysProgFun.MoveNext
    Loop
    
    '��������������
    mrsSysProgFun.Sort = "Id"
    cboFunction.Clear
    cboFunction.addItem ""
    Do While Not mrsSysProgFun.EOF
        cboFunction.addItem mrsSysProgFun!id & "-" & mrsSysProgFun!����
        mrsSysProgFun.MoveNext
    Loop
End Sub

Private Sub cboSystem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '����س�������Ƶ���һ���ؼ�
        cboModule.SetFocus
    End If
End Sub

Private Sub cboUserName_Change()
    'ģ�����ң����������ʾ��listBox��
    '����ģʽ�������Ʋ���
    If cboUserName.Locked Then Exit Sub
    If cboUserName.Text <> "" Then
        lisShowList.Top = cboUserName.Top + cboUserName.Height
        lisShowList.Visible = True
    Else
        lisShowList.Visible = False
        Exit Sub
    End If
    mrsUserName.Filter = "�û��� like '%" & cboUserName.Text & "%'"
    lisShowList.Clear
    With mrsUserName
        lisShowList.Height = 210 * .RecordCount
        If lisShowList.Height > 1470 Then lisShowList.Height = 1470
        Do While Not .EOF
            lisShowList.addItem !�û���
            .MoveNext
        Loop
    End With
End Sub

Private Sub cboUserName_DropDown()
    lisShowList.Visible = False
End Sub

Private Sub cboUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    '���¡��·������,������ת�Ƶ��б���
    If KeyCode = 40 And lisShowList.Visible And lisShowList.ListCount <> 0 Then
        cboUserName.Locked = True
        lisShowList.SetFocus
        lisShowList.ListIndex = 0
    ElseIf KeyCode = 13 Then
        '������س�������Ƶ���һ���ؼ�
        cboSystem.SetFocus
    Else
        KeyCode = 0
    End If
End Sub

Private Sub cboUserName_LostFocus()
    If lisShowList.ListCount = 0 Then lisShowList.Visible = False
End Sub

Private Sub cboWorkStation_Change()
    'ģ�����ң����������ʾ��listBox��
    '����ģʽ�������Ʋ���
    If cboWorkStation.Locked Then Exit Sub
    If cboWorkStation.Text <> "" Then
        lisShowList.Top = cboWorkStation.Top + cboWorkStation.Height
        lisShowList.Visible = True
    Else
        lisShowList.Visible = False
        Exit Sub
    End If
    mrsWorkStation.Filter = "����վ like '%" & cboWorkStation.Text & "%' or ���� like '%" & cboWorkStation.Text & "%'"
    lisShowList.Clear
    With mrsWorkStation
        lisShowList.Height = 210 * .RecordCount
        If lisShowList.Height > 1470 Then lisShowList.Height = 1470
        Do While Not .EOF
            lisShowList.addItem !����վ
            .MoveNext
        Loop
    End With
End Sub

Private Sub cboWorkStation_DropDown()
    lisShowList.Visible = False
End Sub

Private Sub cboWorkStation_KeyDown(KeyCode As Integer, Shift As Integer)
    '���¡��·������,������ת�Ƶ��б���
    If KeyCode = 40 And lisShowList.Visible And lisShowList.ListCount <> 0 Then
        cboWorkStation.Locked = True
        lisShowList.SetFocus
        lisShowList.ListIndex = 0
    ElseIf KeyCode = 13 Then
        '������س�������Ƶ���һ���ؼ�
        cboUserName.SetFocus
    Else
        KeyCode = 0
    End If
End Sub

Private Sub cboWorkStation_LostFocus()
    If lisShowList.ListCount = 0 Then lisShowList.Visible = False
End Sub

Private Sub cmdFind_Click()
    Call FillAuditLog
End Sub

'�����־����
Private Sub FillAuditLog()
    Dim strSQL As String
    Dim lngSystemNo As Long, lngFunctionNo As String
    Dim strModuleNo As String
    Dim i As Long
    
    On Error GoTo errH
    If cboWorkStation.Text <> "" Then strSQL = " And A.����վ = [1]"
    If cboUserName.Text <> "" Then strSQL = strSQL & " And a.�û��� = [2]"
    If cboSystem.Text <> "" Then
        lngSystemNo = Split(cboSystem.Text, "-")(0)
        strSQL = strSQL & " And Nvl(f.ϵͳ,0) = [3]"
    End If
    If cboModule.Text <> "" Then
        strModuleNo = Split(cboModule.Text, "-")(0)
        strSQL = strSQL & " And f.ģ�� = [4]"
    End If
    If cboFunction.Text <> "" Then
        lngFunctionNo = Split(cboFunction.Text, "-")(0)
        strSQL = strSQL & " And f.Id = [5]"
    End If
    strSQL = strSQL & " And a.����ʱ�� between [6] and [7]"
    strSQL = "Select a.�û���, d.����, e.���� ����, a.����վ, a.����ʱ��, Decode(a.��������, 1, '����', 2, '�޸�', 'ɾ��') ��������, 0 ϵͳ, '������������' ����ϵͳ, f.ģ��," & vbNewLine & _
            "       g.���� ����ģ��, f.����, a.��������, a.����˵��" & vbNewLine & _
            "From Zlauditlog A, �ϻ���Ա�� B, ������Ա C, ��Ա�� D, ���ű� E, Zlauditlogconfig F, zlSvrTools G" & vbNewLine & _
            "Where a.�û��� = b.�û��� And b.��Աid = d.Id And b.��Աid = c.��Աid And c.����id = e.Id And c.ȱʡ = 1 And a.��־Id = f.Id And" & vbNewLine & _
            "      f.ģ�� = g.��� And f.ϵͳ Is Null" & strSQL & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.�û���, d.����, e.���� ����, a.����վ, a.����ʱ��, Decode(a.��������, 1, '����', 2, '�޸�', 'ɾ��') ��������, f.ϵͳ, h.���� ����ϵͳ, f.ģ��," & vbNewLine & _
            "       g.���� ����ģ��, f.����, a.��������, a.����˵��" & vbNewLine & _
            "From Zlauditlog A, �ϻ���Ա�� B, ������Ա C, ��Ա�� D, ���ű� E, Zlauditlogconfig F, zlPrograms G, zlSystems H" & vbNewLine & _
            "Where a.�û��� = b.�û��� And b.��Աid = d.Id And b.��Աid = c.��Աid And c.����id = e.Id And c.ȱʡ = 1 And a.��־Id = f.Id And" & vbNewLine & _
            "      f.ģ�� = g.��� And f.ϵͳ = g.ϵͳ And f.ϵͳ = h.���" & strSQL
    frmMDIMain.stbThis.Panels(2).Text = "���ڲ��ң�"
    Set mrsAuditLog = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, cboWorkStation.Text, cboUserName.Text, lngSystemNo, strModuleNo, lngFunctionNo, _
                    CDate(Format(dtpDateStart.value, "YYYY-MM-DD") & " 00:00:00"), CDate(Format(dtpDateEnd.value, "YYYY-MM-DD") & " 23:59:59"))
    With mrsAuditLog
        .Sort = "����ʱ��"
        vsfAuditLogList.Rows = .RecordCount + 1
        For i = 1 To .RecordCount
            vsfAuditLogList.TextMatrix(i, VLL_�û���) = !�û���
            vsfAuditLogList.TextMatrix(i, VLL_��Ա) = !����
            vsfAuditLogList.TextMatrix(i, VLL_����) = !����
            vsfAuditLogList.TextMatrix(i, VLL_����վ) = !����վ
            vsfAuditLogList.TextMatrix(i, VLL_��������) = !��������
            vsfAuditLogList.TextMatrix(i, VLL_ϵͳ���) = !ϵͳ
            vsfAuditLogList.TextMatrix(i, VLL_����ϵͳ) = !����ϵͳ
            vsfAuditLogList.TextMatrix(i, VLL_����ģ����) = !ģ��
            vsfAuditLogList.TextMatrix(i, VLL_����ģ��) = !����ģ��
            vsfAuditLogList.TextMatrix(i, VLL_��������) = !����
            vsfAuditLogList.TextMatrix(i, VLL_����ʱ��) = !����ʱ��
            vsfAuditLogList.TextMatrix(i, VLL_��������) = IIf(Len(!��������) > 50, Mid(!��������, 1, 50) & "...", !��������)
            vsfAuditLogList.TextMatrix(i, VLL_����˵��) = IIf(Len(!����˵��) > 50, Mid(!����˵�� & "", 1, 50) & "...", !����˵�� & "")
            vsfAuditLogList.TextMatrix(i, VLL_��ϸ��������) = !��������
            vsfAuditLogList.TextMatrix(i, VLL_��ϸ����˵��) = !����˵�� & ""
            .MoveNext
        Next
        frmMDIMain.stbThis.Panels(2).Text = "�����ҵ���" & .RecordCount & "�������ݣ�"
        vsfAuditLogList.Tag = frmMDIMain.stbThis.Panels(2).Text
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdLogClear_Click()
    If frmAuditLogClear.ShowMe() Then
        '�жϵ�ǰ�����Ƿ�����־���ݣ�����У�ִ��ˢ�²���������ˢ��
        If vsfAuditLogList.Rows > 1 Then
            Call FillAuditLog
        End If
    End If
End Sub

Private Sub cmdStart_Click()
    On Error GoTo errH
    '������ģ�����־
    Call ExecuteProcedure("Zl_Zlauditlogconfig_Update(" & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID) & ",1)", "����ģ����־")
    vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_״̬���) = 1
    Call RecUpdate(mrsModuleList, "Id = " & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID), "�Ƿ�����", 1)
    vsfAuditLogConfig.Cell(flexcpPicture, vsfAuditLogConfig.Row, VLC_״̬) = imgModuleType.ListImages(2).Picture
    cmdStart.Enabled = False
    CmdStop.Enabled = True
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub CmdStop_Click()
    On Error GoTo errH
    'ͣ�ø�ģ�����־
    Call ExecuteProcedure("Zl_Zlauditlogconfig_Update(" & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID) & ",0)", "ͣ��ģ����־")
    vsfAuditLogConfig.Cell(flexcpPicture, vsfAuditLogConfig.Row, VLC_״̬) = imgModuleType.ListImages(1).Picture
    vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_״̬���) = 0
    Call RecUpdate(mrsModuleList, "Id = " & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID), "�Ƿ�����", 0)
    cmdStart.Enabled = True
    CmdStop.Enabled = False
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub dtpDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '������س�������Ƶ���һ���ؼ�
        cmdFind.SetFocus
    End If
End Sub

Private Sub dtpDateStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '������س�������Ƶ���һ���ؼ�
        dtpDateEnd.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    '��tebControl�ؼ����г�ʼ��
    Call InitTabControl
    
    '���������ݣ���ҪΪ��־���Ҳ�������������
    Call FillBaseData
    
    '���ģ����ͣ����
    Call FillModuleTree
End Sub

'==============================================================================
'=���ܣ� ��ʼTab�ؼ�
'==============================================================================
Private Function InitTabControl() As Boolean
    Dim objTabItem As TabControlItem
On Error GoTo errH
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .OneNoteColors = True
            .DisableLunaColors = True
        End With
        '��һҳ
        Set objTabItem = .InsertItem(0, "��־�鿴", picAuditLogList.hwnd, 0)
        '�ڶ�ҳ
        .InsertItem 1, "��־��ͣ", picAuditLogConfig.hwnd, 0
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With

    InitTabControl = True

    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

'����������
Private Sub FillBaseData()
    Dim rsTmp As ADODB.Recordset
    Dim lngLastSystemNo As Long
    Dim strLastModuleNo As String

    On Error GoTo errH
    '��乤��վ����
    gstrSQL = "Select ����վ,zlspellcode(����վ) ���� From zlClients"
    Set mrsWorkStation = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption)
    Do While Not mrsWorkStation.EOF
        cboWorkStation.addItem mrsWorkStation!����վ
        mrsWorkStation.MoveNext
    Loop
    
    '����û�������
    gstrSQL = "Select �û��� From �ϻ���Ա��"
    Set mrsUserName = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption)
    Do While Not mrsUserName.EOF
        cboUserName.addItem mrsUserName!�û���
        mrsUserName.MoveNext
    Loop
    
    '��¼ϵͳ��ģ�鼰���ܵĶ�Ӧ��ϵ��������ҹ��ܵ�ʹ��
    gstrSQL = "Select a.Id, 0 ϵͳ, '������������' ϵͳ����, a.ģ��, b.���� ģ������, a.����" & vbNewLine & _
            "From Zlauditlogconfig A, zlSvrTools B" & vbNewLine & _
            "Where a.ģ�� = b.��� And a.ϵͳ Is Null" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.Id, a.ϵͳ, b.���� ϵͳ����, a.ģ��, c.���� ģ������, a.����" & vbNewLine & _
            "From Zlauditlogconfig A, zlSystems B, zlPrograms C" & vbNewLine & _
            "Where a.ϵͳ = b.��� And a.ģ�� = c.��� And b.��� = c.ϵͳ"
    Set mrsSysProgFun = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption)
    
    '���Ӧ��ϵͳ�͹���������
    mrsSysProgFun.Sort = "ϵͳ"
    lngLastSystemNo = -1
    cboSystem.addItem ""
    Do While Not mrsSysProgFun.EOF
        If mrsSysProgFun!ϵͳ <> lngLastSystemNo Then
            cboSystem.addItem mrsSysProgFun!ϵͳ & "-" & mrsSysProgFun!ϵͳ����
            lngLastSystemNo = mrsSysProgFun!ϵͳ
        End If
        mrsSysProgFun.MoveNext
    Loop
    
    '������ģ������
    mrsSysProgFun.Sort = "ϵͳ, ģ��"
    cboModule.addItem ""
    Do While Not mrsSysProgFun.EOF
        If mrsSysProgFun!ϵͳ & "-" & mrsSysProgFun!ģ�� <> strLastModuleNo Then
            cboModule.addItem mrsSysProgFun!ģ�� & "-" & mrsSysProgFun!ģ������
            strLastModuleNo = mrsSysProgFun!ϵͳ & "-" & mrsSysProgFun!ģ��
        End If
        mrsSysProgFun.MoveNext
    Loop
    
    '��������������
    mrsSysProgFun.Sort = "Id"
    cboFunction.addItem ""
    Do While Not mrsSysProgFun.EOF
        cboFunction.addItem mrsSysProgFun!id & "-" & mrsSysProgFun!����
        mrsSysProgFun.MoveNext
    Loop
    
    '���ʱ������
    dtpDateStart.value = CurrentDate()
    dtpDateEnd.value = dtpDateStart.value
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'������νṹϵͳ��ģ����Ϣ
Private Sub FillModuleTree()
    Dim lngLastSystemNo As Long   '���һ����ӵ�ϵͳ���
    Dim lngLaseProgNo As Long   '���һ����ӵ�ģ���ģ����
    Dim objNode As Node
    Dim i As Long

    On Error GoTo errH
    gstrSQL = "Select *" & vbNewLine & _
                "From (Select a.Id, 0 ϵͳ, '������������' ϵͳ����, a.ģ��, b.���� ģ������, zlSpellCode(b.����) ����, a.����, a.˵��, a.�Ƿ������, a.�Ƿ�����" & vbNewLine & _
                "       From Zlauditlogconfig A, zlSvrTools B" & vbNewLine & _
                "       Where a.ģ�� = b.��� And a.ϵͳ Is Null" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select a.Id, a.ϵͳ, c.���� ϵͳ����, a.ģ��, b.���� ģ������, zlSpellCode(b.����) ����, a.����, a.˵��, a.�Ƿ������, a.�Ƿ�����" & vbNewLine & _
                "       From Zlauditlogconfig A, zlPrograms B, zlSystems C" & vbNewLine & _
                "       Where a.ϵͳ = b.ϵͳ And a.ģ�� = b.��� And a.ϵͳ = c.���)" & vbNewLine & _
                "Order By ϵͳ, ģ��"
    '������νṹ�������ݣ�˵���Ѿ������˳�ʼ���ˣ������ε���Ϊ���ҹ��ܵ���
    If tvwAuditLogConfig.Nodes.Count = 0 Then
        Set mrsModuleList = CopyNewRec(gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption))
        lngLastSystemNo = -1
        '�������
        With mrsModuleList
            Do While Not .EOF
                If !ϵͳ <> lngLastSystemNo Then
                    Set objNode = tvwAuditLogConfig.Nodes.Add(, , "K_" & !ϵͳ, !ϵͳ����, "system")
                    objNode.Expanded = True
                    lngLastSystemNo = !ϵͳ
                    Set objNode = tvwAuditLogConfig.Nodes.Add("K_" & lngLastSystemNo, tvwChild, "K_" & lngLastSystemNo & "_" & !ģ��, !ģ������, "program")
                    objNode.Tag = !����
                    lngLaseProgNo = !ģ��
                Else
                    If !ģ�� <> lngLaseProgNo Then
                        Set objNode = tvwAuditLogConfig.Nodes.Add("K_" & lngLastSystemNo, tvwChild, "K_" & lngLastSystemNo & "_" & !ģ��, !ģ������, "program")
                        objNode.Tag = !����
                        lngLaseProgNo = !ģ��
                    End If
                End If
                .MoveNext
            Loop
            If .RecordCount <> 0 Then
                tvwAuditLogConfig.Nodes(1).Child.Selected = True
                tvwAuditLogConfig.Tag = tvwAuditLogConfig.SelectedItem.Key
                Call tvwAuditLogConfig_NodeClick(tvwAuditLogConfig.SelectedItem)
            End If
        End With
    Else
        '����궨λ��Ҫ���ҵ���������
        tvwAuditLogConfig.Nodes(tvwAuditLogConfig.Tag).BackColor = &H80000005
        tvwAuditLogConfig.Nodes(tvwAuditLogConfig.Tag).ForeColor = &H80000012
        If mlngCurPos > tvwAuditLogConfig.Nodes.Count Then mlngCurPos = 1
        For i = mlngCurPos To tvwAuditLogConfig.Nodes.Count
            Set objNode = tvwAuditLogConfig.Nodes(i)
            If objNode.Tag <> "" Then
                If objNode.Text Like "*" & txtFind.Text & "*" Or objNode.Tag Like "*" & UCase(txtFind.Text) & "*" Then
                    objNode.Expanded = True
                    objNode.Selected = True
                    objNode.BackColor = &H8000000D
                    objNode.ForeColor = &H80000005
                    tvwAuditLogConfig.Tag = tvwAuditLogConfig.SelectedItem.Key
                    mlngCurPos = i
                    Call tvwAuditLogConfig_NodeClick(objNode)
                    Exit For
                ElseIf i = tvwAuditLogConfig.Nodes.Count Then
                    mlngCurPos = 0
                End If
            End If
        Next
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'���ģ����ͣ��Ϣ
Private Sub FillModuleList()
    Dim i As Long

    On Error GoTo errH
    '����ѯ������Ϣ��䵽������
    With mrsModuleList
        vsfAuditLogConfig.Rows = .RecordCount + 1
        For i = 1 To .RecordCount
            vsfAuditLogConfig.TextMatrix(i, VLC_ID) = !id
            vsfAuditLogConfig.TextMatrix(i, VLC_����ϵͳ) = !ϵͳ����
            vsfAuditLogConfig.TextMatrix(i, VLC_ģ������) = !ģ������
            vsfAuditLogConfig.TextMatrix(i, VLC_��������) = !����
            vsfAuditLogConfig.TextMatrix(i, VLC_˵��) = !˵�� & ""
            If !�Ƿ������ = 1 Then
                vsfAuditLogConfig.TextMatrix(i, VLC_�����) = "��"
            ElseIf !�Ƿ������ = 2 Then
                vsfAuditLogConfig.TextMatrix(i, VLC_�����) = "��"
            Else
                vsfAuditLogConfig.TextMatrix(i, VLC_�����) = ""
            End If
            vsfAuditLogConfig.TextMatrix(i, VLC_״̬���) = !�Ƿ�����
            vsfAuditLogConfig.Cell(flexcpPicture, i, VLC_״̬) = imgModuleType.ListImages(!�Ƿ����� + 1).Picture
            .MoveNext
        Next
        If .RecordCount > 0 Then
            vsfAuditLogConfig.Row = 1
            Call vsfAuditLogConfig_Click
        End If
    End With
    If VScrollVisible(vsfAuditLogConfig) Then
        vsfAuditLogConfig.ColWidth(VLC_˵��) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_ģ������) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                            - vsfAuditLogConfig.ColWidth(VLC_״̬) - vsfAuditLogConfig.ColWidth(VLC_��������) - 350
    Else
        vsfAuditLogConfig.ColWidth(VLC_˵��) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_ģ������) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                            - vsfAuditLogConfig.ColWidth(VLC_״̬) - vsfAuditLogConfig.ColWidth(VLC_��������) - 100
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Private Sub Form_Resize()
    On Error Resume Next
    tbcPage.Left = 0
    tbcPage.Width = Me.ScaleWidth
    tbcPage.Top = 520
    tbcPage.Height = Me.ScaleHeight - tbcPage.Top
    cmdLogClear.Top = 80
    cmdLogClear.Left = Me.ScaleWidth - 200 - cmdLogClear.Width
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsAuditLog = Nothing
    Set mrsModuleList = Nothing
    mlngCurPos = 0
End Sub

Private Sub lisShowList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then
        '�����¡��Ϸ���������·������ʱ��ͬ������cboWorkStation�е�ֵ
        If cboWorkStation.Locked Then
            cboWorkStation.Text = lisShowList.List(lisShowList.ListIndex)
        Else
            cboUserName.Text = lisShowList.List(lisShowList.ListIndex)
        End If
    ElseIf KeyCode = 13 Then
        '������س���ѡ����ǰѡ�е�����
        If cboWorkStation.Locked Then
            lisShowList.Visible = False
            cboWorkStation.Locked = False
            cboWorkStation.SetFocus
        Else
            lisShowList.Visible = False
            cboUserName.Locked = False
            cboUserName.SetFocus
        End If
    End If
End Sub

Private Sub picAuditLogConfig_Resize()
    On Error Resume Next
    vsfAuditLogConfig.Width = picAuditLogConfig.Width - vsfAuditLogConfig.Left - 150
    vsfAuditLogConfig.Height = picAuditLogConfig.Height - vsfAuditLogConfig.Top - 10
    tvwAuditLogConfig.Height = vsfAuditLogConfig.Height
    CmdStop.Left = picAuditLogConfig.Width - CmdStop.Width - 180
    cmdStart.Left = CmdStop.Left - cmdStart.Width
    If VScrollVisible(vsfAuditLogConfig) Then
        vsfAuditLogConfig.ColWidth(VLC_˵��) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_ģ������) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                            - vsfAuditLogConfig.ColWidth(VLC_״̬) - vsfAuditLogConfig.ColWidth(VLC_��������) - 350
    Else
        vsfAuditLogConfig.ColWidth(VLC_˵��) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_ģ������) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                            - vsfAuditLogConfig.ColWidth(VLC_״̬) - vsfAuditLogConfig.ColWidth(VLC_��������) - 100
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub picAuditLogList_Resize()
    On Error Resume Next
    vsfAuditLogList.Width = picAuditLogList.Width - picFind.Width
    vsfAuditLogList.Height = picAuditLogList.Height
    picFind.Left = vsfAuditLogList.Width
    fraDescription.Top = picAuditLogList.Height - fraDescription.Height - 200
    fraDescription.Left = picFind.Left + 150
    fraNote.Top = fraDescription.Top - fraNote.Height - 200
    fraNote.Left = fraDescription.Left
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Caption = "��־�鿴" Then
        cmdLogClear.Visible = True
        frmMDIMain.stbThis.Panels(2).Text = vsfAuditLogList.Tag
    Else
        cmdLogClear.Visible = False
        frmMDIMain.stbThis.Panels(2).Text = ""
    End If
End Sub

Private Sub tvwAuditLogConfig_NodeClick(ByVal Node As MSComctlLib.Node)
    '����Ҳ��б�����
    If tvwAuditLogConfig.Tag <> "" Then
        tvwAuditLogConfig.Nodes(tvwAuditLogConfig.Tag).BackColor = &H80000005
        tvwAuditLogConfig.Nodes(tvwAuditLogConfig.Tag).ForeColor = &H80000012
    End If
    Node.BackColor = &H8000000D
    Node.ForeColor = &H80000005
    tvwAuditLogConfig.Tag = Node.Key
    If tvwAuditLogConfig.SelectedItem.Parent Is Nothing Then
        mrsModuleList.Filter = "ϵͳ = " & Split(tvwAuditLogConfig.SelectedItem.Key, "_")(1)
        vsfAuditLogConfig.MergeCol(VLC_ģ������) = True
    Else
        mrsModuleList.Filter = "ϵͳ = " & Split(tvwAuditLogConfig.SelectedItem.Key, "_")(1) & " And ģ�� = '" & Split(tvwAuditLogConfig.SelectedItem.Key, "_")(2) & "'"
        vsfAuditLogConfig.MergeCol(VLC_ģ������) = False
    End If
    Call FillModuleList
End Sub

Private Sub txtFind_Change()
    mlngCurPos = 0
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String

    If KeyCode = vbKeyReturn Then
        mlngCurPos = mlngCurPos + 1
        Call FillModuleTree
    End If
End Sub

Private Sub vsfAuditLogConfig_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If VScrollVisible(vsfAuditLogConfig) Then
        vsfAuditLogConfig.ColWidth(VLC_˵��) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_ģ������) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                                - vsfAuditLogConfig.ColWidth(VLC_��������) - vsfAuditLogConfig.ColWidth(VLC_״̬) - 350
        vsfAuditLogConfig.ColWidth(VLC_ģ������) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_˵��) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                                - vsfAuditLogConfig.ColWidth(VLC_��������) - vsfAuditLogConfig.ColWidth(VLC_״̬) - 350
        vsfAuditLogConfig.ColWidth(VLC_��������) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_ģ������) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                                - vsfAuditLogConfig.ColWidth(VLC_˵��) - vsfAuditLogConfig.ColWidth(VLC_״̬) - 350
    Else
        vsfAuditLogConfig.ColWidth(VLC_˵��) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_ģ������) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                                - vsfAuditLogConfig.ColWidth(VLC_��������) - vsfAuditLogConfig.ColWidth(VLC_״̬) - 100
        vsfAuditLogConfig.ColWidth(VLC_ģ������) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_˵��) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                                - vsfAuditLogConfig.ColWidth(VLC_��������) - vsfAuditLogConfig.ColWidth(VLC_״̬) - 100
        vsfAuditLogConfig.ColWidth(VLC_��������) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_ģ������) - vsfAuditLogConfig.ColWidth(VLC_�����) _
                                                - vsfAuditLogConfig.ColWidth(VLC_˵��) - vsfAuditLogConfig.ColWidth(VLC_״̬) - 100
    End If
End Sub

Private Sub vsfAuditLogConfig_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = VLC_˵�� Or Col = VLC_����� Then
        Cancel = True
    End If
End Sub

Private Sub vsfAuditLogConfig_Click()
    If vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.RowSel, VLC_״̬���) = 1 Then
        cmdStart.Enabled = False
        CmdStop.Enabled = True
    Else
        cmdStart.Enabled = True
        CmdStop.Enabled = False
    End If
End Sub

Private Sub vsfAuditLogConfig_DblClick()
    On Error GoTo errH
    With vsfAuditLogConfig
        If .MouseRow <> .Row Then Exit Sub
        'ֻ�е����˫����״̬��һ��ʱ���Ž�����ͣ������������������
        If .ColSel = VLC_״̬ Then
            If .TextMatrix(.RowSel, VLC_״̬���) = 1 Then
                Call CmdStop_Click
            Else
                Call cmdStart_Click
            End If
        ElseIf .ColSel = VLC_����� Then
            If .TextMatrix(.RowSel, VLC_�����) = "��" Then
                .TextMatrix(.RowSel, VLC_�����) = "��"
                Call ExecuteProcedure("Zl_Zlauditlogconfig_Update(" & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID) & ",Null,2)", "����Ϊ�������")
            ElseIf .TextMatrix(.RowSel, VLC_�����) = "��" Then
                .TextMatrix(.RowSel, VLC_�����) = "��"
                Call ExecuteProcedure("Zl_Zlauditlogconfig_Update(" & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID) & ",Null,1)", "����Ϊ�����")
            End If
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub vsfAuditLogConfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sinLeft As Single, sinRight As Single
    Dim strTip As String
    
    With vsfAuditLogConfig
        sinLeft = .ColWidth(VLC_ģ������) + .ColWidth(VLC_��������) + .ColWidth(VLC_˵��)
        sinRight = sinLeft + .ColWidth(VLC_�����)
        If X >= sinLeft And X <= sinRight And Y <= 260 Then
            strTip = "ֻ��һЩ�ر���Ҫ�Ĳ������������˲�����" & vbNewLine & _
                           "�������˲��������û�ʹ�ö�Ӧ����ʱ�����û�Ϊ��ͨ����Ա������Ҫ���й���Ա�����֤����д����˵�������û�Ϊϵͳ�����ߣ���ֻ����д����˵����" & vbNewLine & _
                           "���������˲�������������������֤����д����˵����"
            Call ShowTipInfo(.hwnd, strTip, True)
        Else
            Call ShowTipInfo(.hwnd, "")
        End If
    End With
End Sub

Private Sub vsfAuditLogList_Click()
    '�����ÿһ��ʱ�������·���ʾ��ϸ�Ĳ������ݺͲ���˵��
    With vsfAuditLogList
        If .MouseRow <> .Row Or .Row < 1 Then Exit Sub
        '��Ϊ�����ϼ�¼���Ǳ���ȡ��������ݣ���Ҫ����ʾ��ϸ�Ĳ������ݺͲ���˵������ֱ��չʾ�Ѿ����ص���������
        txtOperationContent.Text = .TextMatrix(.Row, VLL_��ϸ��������)
        txtInstructions.Text = .TextMatrix(.Row, VLL_��ϸ����˵��)
    End With
End Sub
