VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPara 
   Caption         =   "��������"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5340
   Icon            =   "frmPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   5340
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   4080
      TabIndex        =   10
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   2880
      TabIndex        =   9
      Top             =   5400
      Width           =   1100
   End
   Begin VB.Frame fraH 
      Height          =   45
      Left            =   0
      TabIndex        =   8
      Top             =   5160
      Width           =   5800
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " �����ϴ�����"
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   5055
      Begin VB.CheckBox chkChooseFMB 
         BackColor       =   &H80000005&
         Caption         =   "�շѵ�"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkChooseFMB 
         BackColor       =   &H80000005&
         Caption         =   "���۵�"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkChooseMB 
         BackColor       =   &H80000005&
         Caption         =   "�շѵ�"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   323
         Width           =   855
      End
      Begin VB.CheckBox chkChooseMB 
         BackColor       =   &H80000005&
         Caption         =   "���۵�"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   323
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����������"
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblchoose1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��������"
         Height          =   180
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView LvwDept 
      Height          =   3285
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5794
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
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   5160
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
            Picture         =   "frmPara.frx":127A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPara.frx":1594
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ���������ң�"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GetPara()
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim strDeptId As String
    Dim intMBType As Integer
    Dim intFMBType As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    
    strDeptId = GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "��������", "")
    intMBType = Val(GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "�������ҵ�������", 0))
    intFMBType = Val(GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "���������ҵ�������", 0))
    
    '��ȡ���������ﲡ�˵Ŀ��ң���Ҫ�ǲ��˾���Ŀ��ң�
    strSql = "Select Distinct ID, ����, ���� " & _
        " From ���ű� A, ��������˵�� B " & _
        " Where a.Id = b.����id And b.�������� In ('�ٴ�', '���', '����', '����', '����', '����', 'Ӫ��') And ������� In (1, 3) And " & _
        " (a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.����ʱ�� Is Null) " & _
        " Order By ���� "
    Set rsTemp = OpenSQLRecord(strSql, "GetDept")
    
    LvwDept.ListItems.Clear
    
    If Not rsTemp.EOF Then
        With rsTemp
            Do While Not .EOF
                LvwDept.ListItems.Add , "D" & !ID, !����, 1, 1
                .MoveNext
            Loop
        End With
    End If
    
    '���ݲ���ֵ���ù�ѡ״̬
    If strDeptId <> "" Then
        With LvwDept
            For i = 1 To .ListItems.Count
                If InStr(1, "," & strDeptId & ",", "," & Mid(.ListItems(i).Key, 2) & ",") > 0 Then
                    .ListItems(i).Checked = True
                End If
            Next
        End With
    End If
    
    '�������ҵ��ݹ���
    If intMBType = 3 Then
        '���е���
        chkChooseMB(0).Value = 1
        chkChooseMB(1).Value = 1
    ElseIf intMBType = 1 Then
        'ֻ���˵���
        chkChooseMB(0).Value = 1
        chkChooseMB(1).Value = 0
    ElseIf intMBType = 2 Then
        'ֻ�շѵ���
        chkChooseMB(0).Value = 0
        chkChooseMB(1).Value = 1
    Else
        '��ѡ�񵥾�
        chkChooseMB(0).Value = 0
        chkChooseMB(1).Value = 0
    End If
    
    '���������ҵ��ݹ���
    If intFMBType = 3 Then
        '���е���
        chkChooseFMB(0).Value = 1
        chkChooseFMB(1).Value = 1
    ElseIf intFMBType = 1 Then
        'ֻ���˵���
        chkChooseFMB(0).Value = 1
        chkChooseFMB(1).Value = 0
    ElseIf intFMBType = 2 Then
        'ֻ�շѵ���
        chkChooseFMB(0).Value = 0
        chkChooseFMB(1).Value = 1
    Else
        '��ѡ�񵥾�
        chkChooseFMB(0).Value = 0
        chkChooseFMB(1).Value = 0
    End If
    
    Exit Sub
errHandle:
    MsgBox "��ȡ������Ϣ����", vbCritical, GSTR_MESSAGE
End Sub

Private Sub cmdSave_Click()
    Dim strDeptId As String
    Dim intMBType As Integer
    Dim intFMBType As Integer
    Dim i As Integer
    
    '�б��й�ѡ�˵Ŀ��ұ���ID��Ϣ
    With LvwDept
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                strDeptId = IIf(strDeptId = "", "", strDeptId & ",") & Mid(.ListItems(i).Key, 2)
            End If
        Next
    End With
    
    '�������ҵ��ݹ���
    If chkChooseMB(0).Value = 1 And chkChooseMB(1).Value = 1 Then
        '���е���
        intMBType = 3
    ElseIf chkChooseMB(0).Value = 1 Then
        'ֻ���˵�
        intMBType = 1
    ElseIf chkChooseMB(1).Value = 1 Then
        'ֻ�շѵ�
        intMBType = 2
    Else
        '��ѡ�񵥾�
        intMBType = 0
    End If
    
    '���������ҵ��ݹ���
    If chkChooseFMB(0).Value = 1 And chkChooseFMB(1).Value = 1 Then
        '���е���
        intFMBType = 3
    ElseIf chkChooseFMB(0).Value = 1 Then
        'ֻ���˵�
        intFMBType = 1
    ElseIf chkChooseFMB(1).Value = 1 Then
        'ֻ�շѵ�
        intFMBType = 2
    Else
        '��ѡ�񵥾�
        intFMBType = 0
    End If
    
    SaveSetting "ZLSOFT", "����ģ��\����ҩ����ҩ��", "��������", strDeptId
    SaveSetting "ZLSOFT", "����ģ��\����ҩ����ҩ��", "�������ҵ�������", intMBType
    SaveSetting "ZLSOFT", "����ģ��\����ҩ����ҩ��", "���������ҵ�������", intFMBType
    
    MsgBox "�����ѱ��棡", vbInformation, ""
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call GetPara
End Sub


