VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmSetExpence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3585
      TabIndex        =   1
      Top             =   4830
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4815
      TabIndex        =   2
      Top             =   4830
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   360
      TabIndex        =   3
      Top             =   4830
      Width           =   1100
   End
   Begin TabDlg.SSTab stab 
      Height          =   4650
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   8202
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "�������"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdBill"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraTitle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "opt(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "opt(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox chk 
         Caption         =   "������ý��н���"
         Height          =   210
         Left            =   225
         TabIndex        =   11
         Top             =   4065
         Width           =   2805
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3585
         Width           =   1695
      End
      Begin VB.OptionButton opt 
         Caption         =   "�����շ�Ʊ��"
         Height          =   210
         Index           =   1
         Left            =   1845
         TabIndex        =   8
         Top             =   555
         Width           =   1530
      End
      Begin VB.OptionButton opt 
         Caption         =   "סԺ����Ʊ��"
         Height          =   210
         Index           =   0
         Left            =   255
         TabIndex        =   7
         Top             =   570
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ع��ý���Ʊ��"
         Height          =   2535
         Left            =   195
         TabIndex        =   5
         Top             =   975
         Width           =   5640
         Begin MSComctlLib.ListView lvwBill 
            Height          =   2220
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   3916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "img16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "������"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "��������"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "���뷶Χ"
               Object.Width           =   2910
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "ʣ��"
               Object.Width           =   1235
            EndProperty
         End
         Begin MSComctlLib.ImageList img16 
            Left            =   735
            Top             =   510
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   1
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSetExpence.frx":0028
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton cmdBill 
         Caption         =   "����Ʊ������(&P)"
         Height          =   350
         Left            =   4095
         TabIndex        =   4
         Top             =   4125
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ĭ�Ͻ��㷽ʽ"
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   3660
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit 'Ҫ���������

Private mblnOK As Boolean
Private mintƱ�� As Integer

Public Function ShowParameter(ByVal frmMain As Object) As Boolean
    
    mblnOK = False
    
    Me.Show 1, frmMain
    
    ShowParameter = mblnOK
    
End Function

Private Sub cmdBill_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL1_BILL_1862", Me)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    '
End Sub

Private Sub cmdOK_Click()
    Dim lngLoop As Long
    
    '�������ע����Ϣ
    
    On Error Resume Next
    
    '���ع��ý���Ʊ��
    If opt(0).Value Then
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 3
    Else
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 1
    End If
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "������ý��н���", chk.Value
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ȱʡ���㷽ʽ", cbo.Text)
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 0
    For lngLoop = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(lngLoop).Checked Then
            SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", Mid(lvwBill.ListItems(lngLoop).Key, 2)
        End If
    Next
            
    mblnOK = True
    Unload Me
End Sub

Private Function ReadBills(ByVal intƱ�� As Integer) As Boolean

    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, blnBill As Boolean
    Dim lngLoop As Long
    Dim strTmp As String
    
    lvwBill.ListItems.Clear
    '��ȡ���ù��ý�������
    gstrSQL = "Select * From Ʊ�����ü�¼ Where Ʊ��=[1] And ʹ�÷�ʽ=2 And ʣ������>0 Order by ʣ������ Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intƱ��)
    blnBill = False
    If Not rsTmp.EOF Then
        For lngLoop = 1 To rsTmp.RecordCount
            Set objItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID, rsTmp!������, , 1)
            objItem.SubItems(1) = Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd")
            objItem.SubItems(2) = rsTmp!��ʼ���� & "," & rsTmp!��ֹ����
            objItem.SubItems(3) = rsTmp!ʣ������
            If rsTmp!ID = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 0) Then
                objItem.Checked = True
                objItem.Selected = True
                blnBill = True
            End If
            rsTmp.MoveNext
        Next
    End If
    If Not blnBill Then SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 0
    
    
    strTmp = Trim(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ȱʡ���㷽ʽ", ""))

    gstrSQL = "SELECT A.���㷽ʽ " & _
                "from ���㷽ʽӦ�� A,���㷽ʽ B where A.���㷽ʽ=B.���� AND A.Ӧ�ó���=[1] AND ���� in (1,2)"
                    
    If intƱ�� = 1 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "�շ�")
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "����")
    End If
    
    cbo.Clear
    
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            cbo.AddItem zlCommFun.NVL(rsTmp("���㷽ʽ").Value)
            If strTmp = zlCommFun.NVL(rsTmp("���㷽ʽ").Value) Then
                cbo.ListIndex = cbo.NewIndex
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
End Function

Private Sub Form_Load()
    
    chk.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "������ý��н���", 1))
    mintƱ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 3))
    If mintƱ�� = 1 Then
        opt(1).Value = True
    Else
        opt(0).Value = True
    End If
    
    Call ReadBills(mintƱ��)
    
End Sub

Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim lngLoop As Long
    
    For lngLoop = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(lngLoop).Key <> Item.Key Then lvwBill.ListItems(lngLoop).Checked = False
    Next
    
    Item.Selected = True
    
End Sub

Private Sub opt_Click(Index As Integer)
    Call ReadBills(IIf(Index = 0, 3, 1))
End Sub
