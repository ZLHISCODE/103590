VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeptSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmDeptSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4635
      TabIndex        =   12
      Top             =   7230
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5895
      TabIndex        =   13
      Top             =   7230
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   225
      TabIndex        =   35
      Top             =   7230
      Width           =   1100
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   2205
      TabIndex        =   34
      Top             =   7260
      Width           =   1335
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6555
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   6735
      Begin VB.Frame fra������Ϣ 
         Caption         =   "������Ϣ"
         Height          =   3975
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   3345
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   840
            MaxLength       =   100
            TabIndex        =   3
            Top             =   1380
            Width           =   1275
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   5
            Left            =   840
            MaxLength       =   50
            TabIndex        =   5
            Top             =   2100
            Width           =   2355
         End
         Begin VB.CommandButton cmd�ϼ� 
            Caption         =   "��"
            Height          =   240
            Left            =   2910
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   2490
            Width           =   255
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   840
            TabIndex        =   1
            Top             =   660
            Width           =   2355
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   840
            MaxLength       =   100
            TabIndex        =   2
            Top             =   1020
            Width           =   1275
         End
         Begin VB.TextBox txtEdit 
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   1
            Left            =   960
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "����"
            Text            =   "111111"
            Top             =   345
            Width           =   1035
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   3540
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cbo������� 
            Height          =   300
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3180
            Width           =   2055
         End
         Begin VB.ComboBox cbo������ 
            Height          =   300
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2820
            Width           =   1815
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   840
            MaxLength       =   3
            TabIndex        =   4
            Top             =   1740
            Width           =   1275
         End
         Begin VB.TextBox txtTemp 
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   840
            MaxLength       =   10
            TabIndex        =   18
            TabStop         =   0   'False
            Tag             =   "����"
            Text            =   "1111111111"
            Top             =   300
            Width           =   1275
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   840
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   6
            Top             =   2460
            Width           =   2355
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&B)"
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   38
            Top             =   1440
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "λ��(&L)"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   2160
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&U)"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&N)"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&S)"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�ϼ�(&P)"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   2520
            Width           =   630
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "Ժ��(&B)"
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   3600
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "�������(&T)"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   3240
            Width           =   1005
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "���Ÿ�����(&D)"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   2880
            Width           =   1170
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "˳��(&R)"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   630
         End
      End
      Begin VB.ComboBox cmb���ƿ�Ŀ���� 
         Height          =   300
         Left            =   3450
         TabIndex        =   11
         Text            =   "cmb���ƿ�Ŀ����"
         Top             =   6165
         Width           =   3105
      End
      Begin VB.Frame fra˵�� 
         Caption         =   "��������˵��"
         Height          =   2310
         Left            =   0
         TabIndex        =   15
         Top             =   4200
         Width           =   3345
         Begin VB.Label lbl˵�� 
            Caption         =   "Label3"
            Height          =   1575
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   3015
         End
      End
      Begin MSComctlLib.ListView lvw���� 
         Height          =   5205
         Left            =   3450
         TabIndex        =   10
         Top             =   480
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   9181
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "��������"
            Object.Tag             =   "��������"
            Text            =   "��������"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "�������"
            Object.Tag             =   "�������"
            Text            =   "�������"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Ĭ��ֵ"
            Object.Tag             =   "Ĭ��ֵ"
            Text            =   "Ĭ��ֵ"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������(&W)"
         Height          =   180
         Left            =   3480
         TabIndex        =   30
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lbl���ƿ�Ŀ���� 
         AutoSize        =   -1  'True
         Caption         =   "�ٴ����ʵ����ƿ��ұ���(&D)"
         Height          =   180
         Left            =   3450
         TabIndex        =   29
         Top             =   5910
         Width           =   2250
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   7470
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
            Picture         =   "frmDeptSet.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptSet.frx":0326
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6465
      Index           =   1
      Left            =   240
      TabIndex        =   31
      Top             =   390
      Visible         =   0   'False
      Width           =   6735
      Begin MSComctlLib.ListView Lvw���� 
         Height          =   5925
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   10451
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   $"frmDeptSet.frx":0640
         Height          =   360
         Left            =   120
         TabIndex        =   33
         Top             =   45
         Width           =   6480
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   6990
      Left            =   120
      TabIndex        =   36
      Top             =   30
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12330
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���Ҳ�����Ӧ"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFind 
      Caption         =   "����(&F)"
      Height          =   255
      Left            =   1500
      TabIndex        =   37
      Top             =   7305
      Width           =   735
   End
   Begin VB.Menu mnuShort 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPatient 
         Caption         =   "���ﲡ��(&O)"
         Index           =   0
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "סԺ����(&I)"
         Index           =   1
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "�����סԺ����(&B)"
         Index           =   2
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "�������ڲ���(&N)"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmDeptSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const mlng���볤�� As Long = 10

Dim mstr�ϼ�����ID As String     '��ǰ�༭���ϼ�����ID
Dim mstrID As String         '��ǰ�༭�Ĳ���ID
Dim mstr�ϼ����� As String    'ԭʼ���ϼ������ֵ
Dim mstr���� As String        'ԭʼ�ı��������ֵ
Dim mint���� As Integer       '�޸�ǰ�����¼����ڵı�����ĳ���
Dim mblnItem As Boolean       '�Ƿ�����ĳ��
Dim mblnChange As Boolean     '�Ƿ�ı���
Dim mblnҩ��  As Boolean
Dim mintԭ�������� As Integer   '0-����ҩ��ҩ�����ʣ�1-��ҩ�����ʣ�2-ֻ��ҩ������
Dim mint���� As Integer         '1-�ٴ�����;2-����;3-�ٴ��Ҳ���
'Dim mint������� As Integer     '1-����;2-סԺ;3-�����סԺ
Dim mint�������_�ٴ� As Integer
Dim mint�������_���� As Integer
Private mlng�������� As Long
Private mstrPrivs As String
Private mint�༭״̬ As Integer     '1-���� 2-�޸�
Private mStr����id As String           '����id
Private mint�༭ģʽ As Integer     '1-�������ʾ 2-��������ʾ
Private mstr�ϼ��� As String
Private mstr���� As String          '��¼��ѡ���ʲô���ʵ�
Private mintInputMethod As Integer   '����¼�뷽ʽ��0-���ϼ����룬1-����¼��
Private mblnPACSInterface As Boolean        '����Ӱ����Ϣϵͳ�ӿ�

Private Function Check��λ״��(ByVal lng����ID As Long, ByVal lng����id As Long, ByVal int���� As Integer) As Boolean
    'int���ʣ�0-�������ʼ��;1-�������Ҷ�Ӧ���
    
    Dim rsTmp As ADODB.Recordset
    
    If lng����ID = 0 Then Exit Function
    
    On Error GoTo ErrHandle
    If int���� = 0 Then
        '�������ʼ�飺�ٴ���������ȡ��ʱ����鴲λ״��
        gstrSQL = "Select 1 From ��λ״����¼ Where (����id = [1] Or ����id = [1]) And Rownum = 1"
    Else
        '�������Ҷ�Ӧ��飺��Ӧ���������ȡ��ʱ����鴲λ��¼
        gstrSQL = "Select 1 From ��λ״����¼ Where (����id = [1] And ����id = [2]) And Rownum = 1"
    End If
        
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��λ״�����", lng����ID, lng����id)
    
    Check��λ״�� = (rsTmp.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub IniStationNo()
    Dim lst As ListItem
    Dim rsRecord As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    lblStationNo.Visible = True
    cmbStationNo.Visible = True
     
    frmDeptSet.Height = frmDeptSet.Height + 50
    tabMain.Height = tabMain.Height + 50
    fraMain(0).Height = fraMain(0).Height + 50
    fra������Ϣ.Height = fra������Ϣ.Height + 50
    fra˵��.Top = fra˵��.Top + 50
    lvw����.Height = lvw����.Height + 250
    fraMain(1).Height = fraMain(1).Height + 50
    Lvw����.Height = Lvw����.Height + 50

    cmdHelp.Top = cmdHelp.Top
    lblFind.Top = cmdHelp.Top + 100
    txtFind.Top = cmdHelp.Top + 25
    cmdOK.Top = cmdHelp.Top
    cmdCancel.Top = cmdHelp.Top
    
    strSQL = "select ���,���� from zlnodelist"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "վ���ѯ")
    
    If rsRecord.RecordCount = 0 Then
        lblStationNo.Visible = False
        cmbStationNo.Visible = False
    Else
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!��� & "-" & rsRecord!����
                rsRecord.MoveNext
            Loop
        End With
    End If

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Ini�������Ҷ�Ӧ(ByVal str����id As String, Optional ByVal int��ʼ As Integer = 1)
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strCon As String
    Dim strCon���� As String
    Dim strCon������� As String
    
    tabMain.Tabs.Clear
    tabMain.Tabs.Add , "_������Ϣ", "������Ϣ"
    
    On Error GoTo ErrHandle
    If glngSys \ 100 <> 1 Then Exit Sub
    
    '��ǰ�������ٴ����������ҷ����������סԺ����ʱ��������
    If mint���� = 0 Or (mint�������_�ٴ� = 0 And mint�������_���� = 0) Then Exit Sub
    
    tabMain.Tabs.Add , "_���Ҳ�����Ӧ", "���Ҳ�����Ӧ"
    
    '��ǰ���ҵķ������ʸı�ʱ�Ÿ��²��������б������
    If (mint�������_�ٴ� = Val(Mid(Lvw����.Tag, 1, 1)) And mint�������_���� = Val(Mid(Lvw����.Tag, 2, 1))) And int��ʼ <> 1 Then Exit Sub
    
    Lvw����.Tag = CStr(mint�������_�ٴ�) & CStr(mint�������_����)
    
    '���ݲ������ʺͷ��������������
    If mint���� = 1 Then
        'ȡ����
        If mint�������_�ٴ� = 1 Then
            strCon���� = " ������� IN(1,3) And �������� = '����' "
        ElseIf mint�������_�ٴ� = 2 Then
            strCon���� = " ������� IN(2,3) And �������� = '����' "
        Else
            strCon���� = " ������� IN(1,2,3) And �������� = '����' "
        End If
    ElseIf mint���� = 2 Then
        'ȡ�ٴ�
        If mint�������_���� = 1 Then
            strCon���� = " ������� IN(1,3) And �������� = '�ٴ�' "
        ElseIf mint�������_���� = 2 Then
            strCon���� = " ������� IN(2,3) And �������� = '�ٴ�' "
        Else
            strCon���� = " ������� IN(1,2,3) And �������� = '�ٴ�' "
        End If
    ElseIf mint���� = 3 Then
        'ȡ�ٴ��Ͳ���
        If mint�������_�ٴ� = 1 Then
            strCon���� = " (������� IN(1,3) And �������� = '����') "
        ElseIf mint�������_�ٴ� = 2 Then
            strCon���� = " (������� IN(2,3) And �������� = '����') "
        Else
            strCon���� = " (������� IN(1,2,3) And �������� = '����') "
        End If
        
        If mint�������_���� = 1 Then
            strCon���� = strCon���� & " Or (������� IN(1,3) And �������� = '�ٴ�') "
        ElseIf mint�������_���� = 2 Then
            strCon���� = strCon���� & " Or (������� IN(2,3) And �������� = '�ٴ�') "
        Else
            strCon���� = strCon���� & " Or (������� IN(1,2,3) And �������� = '�ٴ�') "
        End If
    End If
    
    mstr���� = strCon����
    gstrSQL = " Select Distinct ����||'-'||���� ����,ID From ���ű� " & _
         " Where ID in (Select ����ID From ��������˵�� Where " & strCon���� & ")" & _
         " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
         " Order By ����||'-'||���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ٴ����һ���")
    
    Lvw����.ListItems.Clear
    With rsTmp
        Do While Not .EOF
            Lvw����.ListItems.Add , "_" & !ID, !����, 1, 1
            .MoveNext
        Loop
    End With
    
    'ȡ�������Ҷ�Ӧ��ϵ
    If mstrID <> "" Then
        If mint���� = 1 Then
            strCon = " ����id = [1] "
        ElseIf mint���� = 2 Then
            strCon = " ����id = [1] "
        ElseIf mint���� = 3 Then
            strCon = " ����id = [1] Or ����id = [1] "
        End If
        
        gstrSQL = "Select Distinct ����id,����id From  �������Ҷ�Ӧ Where " & strCon
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�������Ҷ�Ӧ", Val(str����id))
        
        With rsTmp
            Do While Not .EOF
                For n = 1 To Lvw����.ListItems.Count
                    If mint���� = 1 Then
                        If Val(Mid(Lvw����.ListItems(n).Key, 2)) = !����ID Then
                            Lvw����.ListItems(n).Tag = 1
                            Lvw����.ListItems(n).Checked = True
                        End If
                    ElseIf mint���� = 2 Then
                        If Val(Mid(Lvw����.ListItems(n).Key, 2)) = !����ID Then
                            Lvw����.ListItems(n).Tag = 1
                            Lvw����.ListItems(n).Checked = True
                        End If
                    ElseIf mint���� = 3 Then
                        If Val(Mid(Lvw����.ListItems(n).Key, 2)) = !����ID Or Val(Mid(Lvw����.ListItems(n).Key, 2)) = !����ID Then
                            Lvw����.ListItems(n).Tag = 1
                            Lvw����.ListItems(n).Checked = True
                        End If
                    End If
                Next
                .MoveNext
            Loop
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
'    If gstrNodeNo = "-" Then Exit Sub
    
    If cmbStationNo.ListCount = 0 Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub
Private Sub Set�������Ҷ�Ӧ()
    Dim i As Long
    Dim str���� As String
    Dim bln����_�ٴ� As Boolean
    Dim bln����_���� As Boolean
    Dim str�������_�ٴ� As String
    Dim str�������_���� As String
    
    mint���� = 0
    With lvw����
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                str���� = IIF(str���� = "", "", str���� & ",") & .ListItems(i).Text
            End If
        Next
        If InStr(1, str����, "�ٴ�") Then bln����_�ٴ� = True
        If InStr(1, str����, "����") Then bln����_���� = True
        
        If bln����_�ٴ� = True And bln����_���� = True Then
            mint���� = 3
        ElseIf bln����_�ٴ� = True Then
            mint���� = 1
        ElseIf bln����_���� = True Then
            mint���� = 2
        End If
    End With
    
    mint�������_�ٴ� = 0
    mint�������_���� = 0
    With lvw����
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                If InStr(1, .ListItems(i).Text, "�ٴ�") > 0 Then
                    str�������_�ٴ� = IIF(str�������_�ٴ� = "", "", str�������_�ٴ� & ",") & .ListItems(i).SubItems(1)
                End If
                If InStr(1, .ListItems(i).Text, "����") > 0 Then
                    str�������_���� = IIF(str�������_���� = "", "", str�������_���� & ",") & .ListItems(i).SubItems(1)
                End If
            End If
        Next
        
        If InStr(1, str�������_�ٴ�, "�����סԺ����") > 0 Then
            mint�������_�ٴ� = 3
        ElseIf InStr(1, str�������_�ٴ�, "סԺ����") > 0 Then
            mint�������_�ٴ� = 2
        ElseIf InStr(1, str�������_�ٴ�, "���ﲡ��") > 0 Then
            mint�������_�ٴ� = 1
        End If
        
        If InStr(1, str�������_����, "�����סԺ����") > 0 Then
            mint�������_���� = 3
        ElseIf InStr(1, str�������_����, "סԺ����") > 0 Then
            mint�������_���� = 2
        ElseIf InStr(1, str�������_����, "���ﲡ��") > 0 Then
            mint�������_���� = 1
        End If
    End With
    
    Call Ini�������Ҷ�Ӧ(mstrID, 0)
End Sub

Private Sub cbo�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cbo�������.ListIndex = -1
End Sub

Private Sub cmb���ƿ�Ŀ����_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim blnTmp As Boolean
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte
    Dim strTemp As String
    
    Select Case KeyAscii
        Case Is = 13
            '��������ı����ǲ��Ǵ���
    '        If cmb���ƿ�Ŀ����.Enabled Then
    '            If Trim(cmb���ƿ�Ŀ����.Text) = "" Then MsgBox "��ѡ��һ�����ƿ�Ŀ���룡", vbInformation, gstrSysName: Exit Sub
    '            blnTmp = False
    '            For i = 0 To cmb���ƿ�Ŀ����.ListCount - 1
    '                If cmb���ƿ�Ŀ����.List(i) Like Trim(cmb���ƿ�Ŀ����.Text) & "* *" Then
    '                    blnTmp = True
    '                    cmb���ƿ�Ŀ����.ListIndex = i
    '                    Exit For
    '                ElseIf cmb���ƿ�Ŀ����.List(i) Like "* " & Trim(cmb���ƿ�Ŀ����.Text) Then
    '                    blnTmp = True
    '                    cmb���ƿ�Ŀ����.ListIndex = i
    '                    Exit For
    '                ElseIf Trim(cmb���ƿ�Ŀ����.List(i)) = Trim(cmb���ƿ�Ŀ����.Text) Then
    '                    blnTmp = True
    '                    cmb���ƿ�Ŀ����.ListIndex = i
    '                    Exit For
    '                End If
    '            Next
    '            If blnTmp = False Then
    '                MsgBox "��������ƿ�Ŀ���벻���ڣ����������룡", vbExclamation, gstrSysName
    '                cmb���ƿ�Ŀ����.Text = ""
    '                cmb���ƿ�Ŀ����.SetFocus
    '                Exit Sub
    '            End If
    '        End If
            
            strTemp = Trim(UCase(cmb���ƿ�Ŀ����.Text))
            If strTemp = "" Then Exit Sub
            If mStr����id = "" Or cmb���ƿ�Ŀ����.Enabled = False Then
                '�������ٴ�����
                gstrSQL = "Select rownum as id, ����,����  From �ٴ����� where ���� like [1] or ���� like [2] or ���� like [3] Order By ���"
            Else
                gstrSQL = "select rownum as id, A.����,A.����,B.�������� from �ٴ����� A,�ٴ����� B " & _
                    "where A.����=B.��������(+) and b.����ID(+)=[4] and ( a.���� like [1] or a.���� like [2] or a.���� like [3]) order by A.���"
            End If
            
            vRect = zlControl.GetControlRect(cmb���ƿ�Ŀ����.hwnd)
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "���ƿ�Ŀ", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, cmb���ƿ�Ŀ����.Height, True, False, True, strTemp & gstrLike, strTemp & gstrLike, strTemp & gstrLike, Val(mStr����id))
                
            If Not rsTemp Is Nothing Then
                cmb���ƿ�Ŀ����.Text = Format(rsTemp("����"), "!@@@@@") & rsTemp("����")
            End If
    End Select
End Sub

Private Sub cmb���ƿ�Ŀ����_Validate(Cancel As Boolean)
    Dim i As Long
    Dim blnTmp As Boolean
    
    '��������ı����ǲ��Ǵ���
    If cmb���ƿ�Ŀ����.Enabled Then
        If Trim(cmb���ƿ�Ŀ����.Text) = "" Then MsgBox "��ѡ��һ�����ƿ�Ŀ���룡", vbInformation, gstrSysName: Cancel = True: Exit Sub
        blnTmp = False
        For i = 0 To cmb���ƿ�Ŀ����.ListCount - 1
            If cmb���ƿ�Ŀ����.List(i) Like Trim(cmb���ƿ�Ŀ����.Text) & "* *" Then
                blnTmp = True
                cmb���ƿ�Ŀ����.ListIndex = i
                Exit For
            ElseIf cmb���ƿ�Ŀ����.List(i) Like "* " & Trim(cmb���ƿ�Ŀ����.Text) Then
                blnTmp = True
                cmb���ƿ�Ŀ����.ListIndex = i
                Exit For
            ElseIf Trim(cmb���ƿ�Ŀ����.List(i)) = Trim(cmb���ƿ�Ŀ����.Text) Then
                blnTmp = True
                cmb���ƿ�Ŀ����.ListIndex = i
                Exit For
            End If
        Next
        If blnTmp = False Then
            MsgBox "��������ƿ�Ŀ���벻���ڣ����������룡", vbExclamation, gstrSysName
            cmb���ƿ�Ŀ����.Text = ""
            cmb���ƿ�Ŀ����.ListIndex = -1
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim lngTmp As Long
    
    If IsValid() = False Then Exit Sub
        
    '����Ƿ�����ɾ�����ŵ�������ͬ
    If mstrID = "" Then
        If CheckSameDept(txtEdit(2).Text, lngTmp) Then
            If MsgBox("��ǰ¼��Ĳ�����������ɾ���Ĳ���������ͬ��" & vbNewLine & "���ǡ�����ָ������񡿲��ָ�Ҳ�����档", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            mstrID = lngTmp
        End If
    End If
    
    If Save����() = False Then Exit Sub
    
    '�ı������ڵ���ʾ
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    Else
    
    End If
    '��������
    mstrID = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(5).Text = ""
    txtEdit(6).Text = ""
    If mintInputMethod = 0 Then
        '�Զ�����
        txtEdit(1).Text = GetMaxLocalCode(mstr�ϼ�����ID, "���ű�")
    Else
        '����¼�����
        txtEdit(1).Text = ""
    End If
    
    For i = 1 To lvw����.ListItems.Count
        lvw����.ListItems(i).Checked = False
    Next
    lbl���ƿ�Ŀ����.Enabled = False
    cmb���ƿ�Ŀ����.Enabled = False
    cmb���ƿ�Ŀ����.ListIndex = -1
    
    txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ�����ID, "���ű�")
    txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(txtTemp.Text)
    
    Call ShowTab(1)
    txtEdit(2).SetFocus
    
    mblnChange = False
End Sub

Private Function IsValid() As Boolean
    Dim i As Long
    Dim blnTmp As Boolean
    Dim strTemp As String
    Dim int�ֹ������� As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    strSQL = "SELECT ���� FROM ���ű� Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "���㷽ʽ�༭")
    
    txtEdit(6).MaxLength = rsTemp.Fields("����").DefinedSize
    
    For i = 1 To 6
        If i <> 4 Then
            If zlCommFun.StrIsValid(Trim(txtEdit(i).Text), txtEdit(i).MaxLength) = False Then
                Call ShowTab(1)
                txtEdit(i).SetFocus
                zlControl.TxtSelAll txtEdit(i)
                Exit Function
            End If
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)

    If Len(Trim(txtEdit(4).Text)) = 0 And Me.Tag = "�ָ�" Then
        MsgBox "�ϼ�����Ϊ�ա�", vbExclamation, gstrSysName
        Call ShowTab(1)
        txtEdit(4).SetFocus
        Exit Function
    End If
    
    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(1).Text) = 0 Then
            MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
            Call ShowTab(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
    Else
        If Len(txtEdit(1).Text) < txtEdit(1).MaxLength Then
            MsgBox "����ĳ��Ȳ�����", vbExclamation, gstrSysName
            Call ShowTab(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txtEdit(1).Text) Or InStr(txtEdit(1).Text, ",") > 0 Or InStr(txtEdit(1).Text, ".") > 0 Then
        MsgBox "����Ӧ��������ɡ�", vbExclamation, gstrSysName
        Call ShowTab(1)
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        Call ShowTab(1)
        txtEdit(2).Text = ""
        txtEdit(2).SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtEdit(2).Text, vbFromUnicode)) > 100 Then
        MsgBox "���Ƴ��Ȳ��ܳ���50�����ֻ���100���ַ���������¼�룡", vbInformation, gstrSysName
        txtEdit(2).SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtEdit(3).Text, vbFromUnicode)) > 100 Then
        MsgBox "���볤�Ȳ��ܳ���100���ַ���������¼�룡", vbInformation, gstrSysName
        txtEdit(3).SetFocus
        Exit Function
    End If
    
    '���¼�������ı����ǲ��Ǵ���
    If glngSys \ 100 = 8 Then
        'ҩ��ϵͳ������
    Else
        '���ٴ����ʵ��ж�
        If cmb���ƿ�Ŀ����.Enabled Then
            If Trim(cmb���ƿ�Ŀ����.Text) = "" Then
                Call ShowTab(1)
                MsgBox "��Ϊ�ٴ����������������ƿ�Ŀ���롣", vbExclamation, gstrSysName
                cmb���ƿ�Ŀ����.SetFocus
                Exit Function
            End If
            blnTmp = False
            For i = 0 To cmb���ƿ�Ŀ����.ListCount - 1
                If cmb���ƿ�Ŀ����.List(i) = cmb���ƿ�Ŀ����.Text Then
                    blnTmp = True
                    Exit For
                End If
            Next
            If blnTmp = False Then
                MsgBox "��������ƿ�Ŀ���벻���ڣ����������룡", vbExclamation, gstrSysName
                Call ShowTab(1)
                cmb���ƿ�Ŀ����.Text = ""
                cmb���ƿ�Ŀ����.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '��鲿�ŵĹ������ʱ仯����Ҫ��ҩ��ҩ�����ʱ仯���������ҩ��ҩ�����ʵı任�����棬�п������ʾ
    On Error Resume Next
    If mstrID <> "" Then
        int�ֹ������� = 0
        For i = 1 To lvw����.ListItems.Count
            If lvw����.ListItems(i).Checked = True Then
                If int�ֹ������� <> 1 Then
                    If InStr(lvw����.ListItems(i), "ҩ��") > 0 Or lvw����.ListItems(i) = "�Ƽ���" Then
                        int�ֹ������� = 1
                    ElseIf InStr(lvw����.ListItems(i), "ҩ��") > 0 Then
                        int�ֹ������� = 2
                    End If
                End If
            End If
        Next
        If int�ֹ������� <> mintԭ�������� Then
            gstrSQL = "select 1 from ҩƷ��� where �ⷿID=[1] and rownum=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
            
            If rsTemp.RecordCount > 0 Then
                If MsgBox("�ò��ź��е�ҩ���ҩ�����ʷ����˱仯�����ܻ�Ӱ����ҩƷ�ķ������ԡ��Ƿ�ȷ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    IsValid = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save����() As Boolean
    Dim strSQL As String
    Dim lng����ID As Long
    Dim str�������� As String
    Dim str�ٴ����� As String
    Dim i As Integer, int���� As Integer
    Dim nod As Node
    Dim lst As ListItem
    Dim str�������� As String
    Dim strվ�� As String
    Dim BeginTrans As Boolean
    Dim arrSQL As Variant
    Dim lngType As Long
    
    On Error GoTo ErrHandle
    
    arrSQL = Array()
    '������ѡ�еĹ�����������һ����
    For i = 1 To lvw����.ListItems.Count
        If lvw����.ListItems(i).Checked = True Then
            str�������� = str�������� & lvw����.ListItems(i) & ":"
            
            If mblnҩ�� = True Then
                'ҩ��ֻ�������ﲡ��
                str�������� = str�������� & "1:"
            Else
                Select Case lvw����.ListItems(i).SubItems(1)
                     Case "���ﲡ��"
                        str�������� = str�������� & "1:"
                     Case "סԺ����"
                        str�������� = str�������� & "2:"
                     Case "�����סԺ����"
                        str�������� = str�������� & "3:"
                     Case Else
                        str�������� = str�������� & "0:"
                End Select
            End If
        End If
    Next
    If cmb���ƿ�Ŀ����.Enabled = False Then
        str�ٴ����� = ""
    Else
        str�ٴ����� = Trim(Left(cmb���ƿ�Ŀ����.List(cmb���ƿ�Ŀ����.ListIndex), 4))
    End If
    
    If mstrID = "" Then       '����һ����¼
        If Check�ظ�����(mstr�ϼ�����ID, Trim(txtEdit(2).Text)) = True Then
            MsgBox "�ü��������иò��ţ����������ͬ���ţ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        lng����ID = Sys.NextId("���ű�")
        lngType = RISBaseItemOper.AddNew
        gstrSQL = "zl_���ű�_insert(" & lng����ID & "," & IIF(mstr�ϼ�����ID = "", "null", mstr�ϼ�����ID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & UCase(txtEdit(3).Text) & "','" & txtEdit(5).Text & "','" & str�������� & "','" & str�ٴ����� & "' "
        
        If cmbStationNo.Text = "" Then
            strվ�� = "Null"
        Else
            strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
        End If
        
        gstrSQL = gstrSQL & ",'" & Trim(cbo�������.Text) & "'," & IIF(cmbStationNo.Text = "", "Null", strվ��)
        gstrSQL = gstrSQL & "," & IIF(txtEdit(0).Text = "", "Null", txtEdit(0).Text)
        gstrSQL = gstrSQL & ",'" & txtEdit(6).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        '�޸������������
        With frmDeptManage.tvwMain_S
            '���ӵ�TreeView��
            If mint�༭ģʽ = 1 Then
                Set nod = .Nodes.Add(IIF(mstr�ϼ�����ID = "", "Root", "C" & mstr�ϼ�����ID), tvwChild, _
                    "C" & lng����ID, "��" & txtTemp.Text & txtEdit(1).Text & "��" & txtEdit(2).Text, "Dept", "Dept")
                nod.Sorted = True
            Else
                Set nod = .Nodes.Add(IIF(mstr�ϼ�����ID = "", "Root", "C" & mstr�ϼ���), tvwChild, _
                    "C" & txtEdit(1).Text & "|" & lng����ID, "��" & txtTemp.Text & txtEdit(1).Text & "��" & txtEdit(2).Text, "Dept", "Dept")
                nod.Sorted = True
            End If
            '���ӵ�ListView��
        End With
        With frmDeptManage.lvwMain
            If frmDeptManage.tvwMain_S.SelectedItem.Key = IIF(mstr�ϼ�����ID = "", "Root", "C" & mstr�ϼ�����ID) Then
                Set lst = .ListItems.Add(, "C" & lng����ID, txtEdit(2).Text, "Dept", "Dept")
                For i = 2 To .ColumnHeaders.Count
                    Select Case .ColumnHeaders(i).Text
                        Case "����"
                            lst.SubItems(i - 1) = txtTemp.Text & txtEdit(1).Text
                        Case "����"
                            lst.SubItems(i - 1) = txtEdit(2).Text
                        Case "����"
                            lst.SubItems(i - 1) = txtEdit(3).Text
                        Case "λ��"
                            lst.SubItems(i - 1) = txtEdit(5).Text
                        Case "����ʱ��"
                            lst.SubItems(i - 1) = Format(Sys.Currentdate, "yyyy-MM-dd")
                        Case "����ʱ��"
                            lst.SubItems(i - 1) = "3000-01-01"
                        Case "�ϼ�����"
                            lst.SubItems(i - 1) = txtEdit(4).Text
                    End Select
                Next
                If .ListItems.Count = 1 Then
                    .ListItems(1).Selected = True
                    Call frmDeptManage.lvwMain_ItemClick(.ListItems(1))
                End If
            End If
        End With
        
    Else
        '�޸�
        lng����ID = Val(mstrID)
        lngType = RISBaseItemOper.Modify
        gstrSQL = "zl_���ű�_update(" & mstrID & "," & IIF(mstr�ϼ�����ID = "", "null", mstr�ϼ�����ID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & UCase(txtEdit(3).Text) & "','" & txtEdit(5).Text & "'," & Len(mstr����) + 1 & ",'" & str�������� & "','" & str�ٴ����� & "' "
        
        gstrSQL = gstrSQL & ",'" & Trim(cbo�������.Text) & "',"
        If Me.cbo������.ListIndex <= 0 Then
            gstrSQL = gstrSQL & "null,"
        Else
            gstrSQL = gstrSQL & Me.cbo������.ItemData(Me.cbo������.ListIndex) & ","
        End If
        If cmbStationNo.Text = "" Then
            strվ�� = "Null"
        Else
            strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
        End If
        
        gstrSQL = gstrSQL & IIF(cmbStationNo.Text = "", "Null", strվ��)
        gstrSQL = gstrSQL & "," & IIF(txtEdit(0).Text = "", "Null", txtEdit(0).Text)
        gstrSQL = gstrSQL & ",'" & txtEdit(6).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    If glngSys \ 100 = 1 Then
        With Lvw����
            For i = 1 To .ListItems.Count
                If .ListItems(i).Checked = True Then
                    str�������� = IIF(str�������� = "", "", str�������� & ",") & Mid(.ListItems(i).Key, 2)
                End If
            Next
        End With
        gstrSQL = "Zl_�������Ҷ�Ӧ_Update(" & IIF(mstrID <> "", Val(mstrID), lng����ID) & "," & mint���� & "," & IIF(str�������� = "", "Null", "'" & str�������� & "'") & ")"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    gcnOracle.BeginTrans: BeginTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
    Next
    
    If glngSys \ 100 = 1 And mblnPACSInterface Then
        If Not gobjRIS Is Nothing Then
            If gobjRIS.HISBasicDictTable(10, lngType, lng����ID) <> 1 Then
                gcnOracle.RollbackTrans
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISBasicDictTable)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISBasicDictTable)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    gcnOracle.CommitTrans: BeginTrans = False

    Call frmDeptManage.FillTree
    Save���� = True
    Exit Function
ErrHandle:
    If BeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ChangeCode(nod As Node, ByVal strOldCode As String, ByVal strNewCode As String)
'����:�ı��¼��ı�������
    Dim nodChild As Node
    
    Set nodChild = nod.Child
    Do Until nodChild Is Nothing
        nodChild.Text = strNewCode & Mid(nodChild.Text, Len(strOldCode))
        ChangeCode nodChild, strOldCode, strNewCode
        Set nodChild = nodChild.Next
    Loop
End Sub

Public Sub �༭����(ByVal strPrivs As String, strID As String, ByVal int�༭״̬ As Integer, ByVal int�༭ģʽ As Integer, ByVal str�ϼ����� As String, Optional str�ϼ�ID As String)
'    On Error GoTo errHandle
    
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim str���� As String
    Dim str�������_�ٴ� As String
    Dim str�������_���� As String
    
    mstrPrivs = strPrivs
    mint�༭״̬ = int�༭״̬
    mStr����id = strID
    mint�༭ģʽ = int�༭ģʽ
    mstr�ϼ��� = str�ϼ�����
    mstr�ϼ�����ID = str�ϼ�ID
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockReadOnly
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    mint���� = 0
    
    mintInputMethod = IIF(Val(zlDatabase.GetPara("����¼�����", glngSys, 1001, "0")) = 1, 1, 0)
    tabMain.Tabs.Clear
    tabMain.Tabs.Add , "_������Ϣ", "������Ϣ"
    
    mlng�������� = Val(zlDatabase.GetPara("��������", glngSys, 0))
        
    Call IniStationNo
    
    Call Cbo.SetListHeight(cmb���ƿ�Ŀ����, cmb���ƿ�Ŀ����.Height * 16)
    
    mblnҩ�� = (glngSys \ 100 = 8)
    If mblnҩ�� = True Then
        'ҩ��ϵͳ��Ҫ���⴦��
        lbl���ƿ�Ŀ����.Visible = False
        cmb���ƿ�Ŀ����.Visible = False
        
'        lbl��������.Top = lbl���ƿ�Ŀ����.Top
'        lvw����.Top = cmb���ƿ�Ŀ����.Top
        lvw����.Height = fra������Ϣ.Height + fra˵��.Height
        
        lvw����.ColumnHeaders(2).Text = "˵��"
    Else
        lbl���ƿ�Ŀ����.Top = lvw����.Top + lvw����.Height + 40
        cmb���ƿ�Ŀ����.Top = lbl���ƿ�Ŀ����.Top + lbl���ƿ�Ŀ����.Height + 50
        lvw����.ToolTipText = "������ѡ��ʱ˫���򰴡�C�����ɸı�������"
    End If
    mstrID = strID
    '���Ÿ�����ѡ��
    gstrSQL = "select a.id,'��'||a.���||'��'||a.���� ���� from ��Ա�� a, ������Ա b where a.id=b.��Աid and b.����id=[1] order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrID)
    With cbo������
        .Clear
        .AddItem ""
        .ItemData(0) = -1
        For i = 1 To rsTemp.RecordCount
            .AddItem NVL(rsTemp!����)
            .ItemData(i) = rsTemp!ID
            rsTemp.MoveNext
        Next
    End With
    
    cbo�������.ListIndex = -1
    If strID <> "" Then
        gstrSQL = "select A.����,A.����,A.����,A.λ��,A.�������,B.���� as �ϼ�����,B.���� as �ϼ�����,B.ID as �ϼ�ID,A.վ�� " _
                & ",A.���Ÿ�����,A.˳��,A.���� " _
                & "from ���ű� A,���ű� B  where A.�ϼ�ID=B.ID(+) and A.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
        mstr�ϼ�����ID = IIF(IsNull(rsTemp("�ϼ�ID")), "", rsTemp("�ϼ�ID"))
        mstr�ϼ����� = IIF(IsNull(rsTemp("�ϼ�����")), "", rsTemp("�ϼ�����"))
        
        
        If Mid(mstr�ϼ�����, 1, 1) = "-" Then
            '�ָ���ɾ������
            Me.Tag = "�ָ�"
            mstr�ϼ����� = ""
            mstr�ϼ�����ID = ""
            txtEdit(4).Text = ""
        Else
            txtEdit(4).Text = IIF(IsNull(rsTemp("�ϼ�����")), "��", rsTemp("�ϼ�����"))
            If mintInputMethod = 0 Then '�����¼�
                txtTemp.Text = mstr�ϼ�����
                'ȡ���ϼ����룬�������볤�ȵ�ֵ
                txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ�����ID, "���ű�")
                'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
                txtEdit(1).Text = Mid(rsTemp("����"), Len(txtTemp.Text) + 1)
                mstr���� = rsTemp("����")
                '��������ӽڵ����ڵ������
                mint���� = GetDownCodeLength(mstrID, "���ű�")
                '10 - (mint���� - Len(mstr����))�����ʽ����˼��ҪΪ���ĺ��ӵı����������
                txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10 - (mint���� - Len(mstr����)), txtTemp.MaxLength) - Len(mstr�ϼ�����)
            Else
                txtTemp.Text = ""
                txtTemp.MaxLength = 0
                txtEdit(1).Text = rsTemp("����")
                mstr���� = rsTemp("����")
                txtEdit(1).MaxLength = 10
            End If
        End If
        
        txtEdit(2).Text = rsTemp("����")
        txtEdit(3).Text = IIF(IsNull(rsTemp("����")), "", rsTemp("����"))
        txtEdit(5).Text = IIF(IsNull(rsTemp("λ��")), "", rsTemp("λ��"))
        txtEdit(0).Text = IIF(IsNull(rsTemp("˳��")), "", rsTemp("˳��"))
        txtEdit(6).Text = IIF(IsNull(rsTemp("����")), "", rsTemp("����"))
        
        With cbo�������
            For i = 0 To .ListCount - 1
                If .List(i) = NVL(rsTemp!�������) Then
                    .ListIndex = i: Exit For
                End If
            Next
            If Trim(NVL(rsTemp!�������)) <> "" And .ListIndex < 0 Then
                .AddItem NVL(rsTemp!�������): .ListIndex = .NewIndex
            End If
        End With
        With cbo������
            For i = 0 To .ListCount - 1
                If .ItemData(i) = NVL(rsTemp!���Ÿ�����) Then
                    .ListIndex = i: Exit For
                End If
            Next
        End With
        
        SetStationNo (IIF(IsNull(rsTemp("վ��")), "", rsTemp("վ��")))
    Else
        If mintInputMethod = 0 Then '�����¼�
            If str�ϼ�ID = "oot" Then
                mstr�ϼ�����ID = ""
                mstr�ϼ����� = ""
                txtTemp.Text = ""
                txtEdit(4).Text = "��"
                'ȡ���ϼ����룬�������볤�ȵ�ֵ
                txtTemp.MaxLength = GetLocalCodeLength("", "���ű�")
            Else
                gstrSQL = "select ���� as �ϼ�����,���� as �ϼ�����,ID as �ϼ�ID from ���ű� where ID=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str�ϼ�ID))
                            
                mstr�ϼ�����ID = IIF(IsNull(rsTemp("�ϼ�ID")), "", rsTemp("�ϼ�ID"))
                mstr�ϼ����� = IIF(IsNull(rsTemp("�ϼ�����")), "", rsTemp("�ϼ�����"))
                txtEdit(4).Text = IIF(IsNull(rsTemp("�ϼ�����")), "��", rsTemp("�ϼ�����"))
                txtTemp.Text = mstr�ϼ�����
                '�жϱ����Ƿ�����
                If Len(mstr�ϼ�����) = mlng���볤�� Then
                    MsgBox "�����������Ӳ����ˣ����볤���Ѿ��þ���", vbExclamation, gstrSysName
                    Exit Sub
                End If
                'ȡ���ϼ����룬�������볤�ȵ�ֵ
                txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ�����ID, "���ű�")
                
                'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
            End If
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(mstr�ϼ�����)
            txtEdit(1).Text = GetMaxLocalCode(mstr�ϼ�����ID, "���ű�")
            mstr���� = mstr�ϼ����� & txtEdit(1).Text
        Else
            '����¼�����
            If str�ϼ�ID = "oot" Then
                mstr�ϼ�����ID = ""
                mstr�ϼ����� = ""
                txtTemp.Text = ""
                txtEdit(4).Text = "��"
                'ȡ���ϼ����룬�������볤�ȵ�ֵ
                txtTemp.MaxLength = 0
            Else
                gstrSQL = "select ���� as �ϼ�����,���� as �ϼ�����,ID as �ϼ�ID from ���ű� where ID=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str�ϼ�ID))
                            
                mstr�ϼ�����ID = IIF(IsNull(rsTemp("�ϼ�ID")), "", rsTemp("�ϼ�ID"))
                mstr�ϼ����� = IIF(IsNull(rsTemp("�ϼ�����")), "", rsTemp("�ϼ�����"))
                txtEdit(4).Text = IIF(IsNull(rsTemp("�ϼ�����")), "��", rsTemp("�ϼ�����"))
                txtTemp.Text = ""
                'ȡ���ϼ����룬�������볤�ȵ�ֵ
                txtTemp.MaxLength = 0
                
                'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
            End If
            txtEdit(1).MaxLength = 10
            txtEdit(1).Text = ""
            mstr���� = mstr�ϼ����� & txtEdit(1).Text
        End If
    End If
    '��ʾ��������
    If rsTemp.State = 1 Then rsTemp.Close
    
    If strID = "" Then
        gstrSQL = "select ����,������ as ȱʡ����,˵��,null as ��������,null as ������� from �������ʷ��� order by decode(��������,null,1,0) ,����"
    Else
        gstrSQL = "select A.����,A.������ as ȱʡ����,A.˵��,B.��������,B.������� from �������ʷ��� A,��������˵�� B where A.����=B.��������(+) and b.����ID(+)=[1] order by decode(��������,null,1,0),A.����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    lbl���ƿ�Ŀ����.Enabled = False
    cmb���ƿ�Ŀ����.Enabled = False
        
    If rsTemp.EOF Then
        mblnChange = False
        frmDeptSet.Show vbModal
        Exit Sub
    End If
        
    Dim lst As ListItem
    Do Until rsTemp.EOF
        If InStr(1, mstrPrivs, ";" & "������������ⷿ" & ";") = 0 And rsTemp("����") = "����ⷿ" Then
            rsTemp.MoveNext
        Else
            Select Case IIF(IsNull(rsTemp("��������")), rsTemp("ȱʡ����"), rsTemp("�������"))
                 Case 1
                    strTemp = "���ﲡ��"
                 Case 2
                    strTemp = "סԺ����"
                 Case 3
                    strTemp = "�����סԺ����"
                 Case Else
                    strTemp = "�������ڲ���"
            End Select
            Set lst = lvw����.ListItems.Add(, rsTemp("����"), rsTemp("����"))
            If mblnҩ�� = True Then
                lst.SubItems(1) = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
            Else
                lst.SubItems(1) = strTemp
            End If
            lst.ListSubItems(1).Tag = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
            lst.Tag = rsTemp("ȱʡ����")
            If Not IsNull(rsTemp("��������")) Then
                lst.SubItems(2) = 1
                lst.Checked = True
                If lst.Text = "�ٴ�" Then
                    lbl���ƿ�Ŀ����.Enabled = True
                    cmb���ƿ�Ŀ����.Enabled = True
                End If
            End If
            
            str���� = IIF(str���� = "", "", str���� & ",") & rsTemp("��������")
            If rsTemp("��������") = "�ٴ�" Then
                str�������_�ٴ� = IIF(str�������_�ٴ� = "", "", str�������_�ٴ� & ",") & strTemp
            End If
            
            If rsTemp("��������") = "����" Then
                str�������_���� = IIF(str�������_���� = "", "", str�������_���� & ",") & strTemp
            End If
        
            rsTemp.MoveNext
        End If
    Loop
    
    '��¼��ʼ�Ĳ������ʺͷ������
    mint���� = 0
    mint�������_�ٴ� = 0
    mint�������_���� = 0
    
    If InStr(1, str����, "�ٴ�") > 0 And InStr(1, str����, "����") > 0 Then
        mint���� = 3
    ElseIf InStr(1, str����, "����") > 0 Then
        mint���� = 2
    ElseIf InStr(1, str����, "�ٴ�") > 0 Then
        mint���� = 1
    End If
    
    If InStr(1, str�������_�ٴ�, "�����סԺ����") > 0 Then
        mint�������_�ٴ� = 3
    ElseIf InStr(1, str�������_�ٴ�, "סԺ����") > 0 Then
        mint�������_�ٴ� = 2
    ElseIf InStr(1, str�������_�ٴ�, "���ﲡ��") > 0 Then
        mint�������_�ٴ� = 1
    End If
    
    If InStr(1, str�������_����, "�����סԺ����") > 0 Then
        mint�������_���� = 3
    ElseIf InStr(1, str�������_����, "סԺ����") > 0 Then
        mint�������_���� = 2
    ElseIf InStr(1, str�������_����, "���ﲡ��") > 0 Then
        mint�������_���� = 1
    End If
    
    Lvw����.Tag = CStr(mint�������_�ٴ�) & CStr(mint�������_����)
    
    lvw����.ListItems(1).Selected = True
    lvw����_ItemClick lvw����.ListItems(1)
    '��ʾ�����ٴ�����
    If rsTemp.State = 1 Then rsTemp.Close
    
    If strID = "" Or cmb���ƿ�Ŀ����.Enabled = False Then
        '�������ٴ�����
        gstrSQL = "select ����,����,null as �������� from �ٴ����� order by ���"
    Else
        gstrSQL = "select A.����,A.����,B.�������� from �ٴ����� A,�ٴ����� B " & _
            "where A.����=B.��������(+) and b.����ID(+)=[1] order by A.���"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    cmb���ƿ�Ŀ����.Clear
    Do Until rsTemp.EOF
        cmb���ƿ�Ŀ����.AddItem Format(rsTemp("����"), "!@@@@@") & rsTemp("����")
        If Not IsNull(rsTemp("��������")) Then
            cmb���ƿ�Ŀ����.ListIndex = cmb���ƿ�Ŀ����.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    
    '��¼ԭ���Ĺ�������
    mintԭ�������� = 0
    For i = 1 To lvw����.ListItems.Count
        If lvw����.ListItems(i).Checked = True Then
            If mintԭ�������� <> 1 Then
                If InStr(lvw����.ListItems(i), "ҩ��") > 0 Or lvw����.ListItems(i) = "�Ƽ���" Then
                    mintԭ�������� = 1
                ElseIf InStr(lvw����.ListItems(i), "ҩ��") > 0 Then
                    mintԭ�������� = 2
                End If
            End If
        End If
    Next
    
    '���Ҳ�������
    Call Ini�������Ҷ�Ӧ(mstrID, 1)
    
    '��ɳ�ʼ��
    If rsTemp.State = 1 Then rsTemp.Close
    
    mblnChange = False
    frmDeptSet.Show vbModal
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd�ϼ�_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim str���� As String
    Dim int����  As Integer
    
    If mstrID <> "" Then
        strSQL = "select id,�ϼ�id,����,����,���� from ���ű� where ����ʱ��=to_date('3000-01-01','YYYY-MM-DD') and id<>" & mstrID & " start with �ϼ�id is null connect by prior id =�ϼ�id And �ϼ�id<>" & mstrID
    Else
        strSQL = "select id,�ϼ�id,����,����,���� from ���ű� where ����ʱ��=to_date('3000-01-01','YYYY-MM-DD') start with �ϼ�id is null connect by prior id =�ϼ�id "
    End If
    strID = mstr�ϼ�����ID
    str���� = txtEdit(4).Text
    str���� = txtTemp.Text
    blnRe = frmTreeSel.ShowTree(strSQL, strID, str����, str����, mstrID, "���ű�", "���в���", , mstr����, 0, 0, 0, False)
    '�ɹ�����
    If blnRe Then       '�µı����Ŀ��
        int���� = GetLocalCodeLength(strID, "���ű�")
        'ֻ���޸Ĳ��б�Ҫ���
        If mstrID <> "" Then
            If mint���� - Len(mstr����) + IIF(int���� = 0, Len(str����) + 1, int����) > 10 Then
                MsgBox "����ϼ������ʣ���Ϊ���ı���̫���ˡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        mstr�ϼ�����ID = strID
        txtEdit(4).Text = str����
        txtTemp.MaxLength = int����
        txtTemp.Text = str����
        If mstrID <> "" Then
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10 - (mint���� - Len(mstr����)), txtTemp.MaxLength) - Len(str����)
        Else
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(str����)
        End If
        txtEdit(1).Text = GetMaxLocalCode(mstr�ϼ�����ID, "���ű�")
    End If
    mblnChange = True
    '���ü��²���˳��
    If CheckOrder = True Then
        txtEdit(0).SetFocus
    End If
End Sub

Private Sub Form_Activate()
    txtEdit(2).SetFocus
    lbl˵��.Move 130, 260, fra˵��.Width - 160, fra˵��.Height - 400
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    '���˺�:2008/03/18����
    On Error GoTo ErrHandle
    gstrSQL = "Select ����,���� From ���Ż������ order by ����"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With cbo�������
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!����)
            rsTemp.MoveNext
        Loop
    End With
    If mint�༭״̬ = 1 Then
        cbo������.Enabled = False
    Else
        cbo������.Enabled = True
    End If
    
    lblFind.Visible = False
    txtFind.Visible = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Lvw����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = False And Val(Lvw����.ListItems(Item.Index).Tag) = 1 Then
        If Check��λ״��(Val(mstrID), Val(Mid(Item.Key, 2)), 1) Then
            MsgBox "�ò��Ŵ��ڴ�λ��¼������ȡ����Ӧ��ϵ��", vbInformation, gstrSysName
            Item.Checked = True
            Exit Sub
        End If
    End If
End Sub


Private Sub lvw����_DblClick()
    If mblnItem = False Then Exit Sub
    Call ChangeServer
    Call Set�������Ҷ�Ӧ
    mblnItem = False
End Sub

Private Sub ChangeServer()
    If lvw����.SelectedItem Is Nothing Then Exit Sub
    If mblnҩ�� = True Then Exit Sub
    
    With lvw����.SelectedItem
        If .Checked = False Then Exit Sub
        Select Case .SubItems(1)
             Case "���ﲡ��"
                .SubItems(1) = "סԺ����"
             Case "סԺ����"
                .SubItems(1) = "�����סԺ����"
             Case "�����סԺ����"
                .SubItems(1) = "�������ڲ���"
             Case Else
                If .Tag <> 0 Then .SubItems(1) = "���ﲡ��"
        End Select
        mblnChange = True
    End With
End Sub

Private Sub lvw����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    Dim bln����ⷿ As Boolean
    
    mblnChange = True
    
    If Item.Text = "�ٴ�" Or Item.Text = "����" Then
        If Item.Checked = False And Val(Item.SubItems(2)) = 1 Then
            If Check��λ״��(Val(mstrID), 0, 0) Then
                MsgBox "�ò��Ŵ��ڴ�λ��¼������ȡ�������ʣ�", vbInformation, gstrSysName
                Item.Checked = True
                Exit Sub
            End If
        End If
    End If
    
    If Item.Text = "�ٴ�" Then
        If Item.Checked = False Then
            lbl���ƿ�Ŀ����.Enabled = False
            cmb���ƿ�Ŀ����.Enabled = False
            cmb���ƿ�Ŀ����.ListIndex = -1
        Else
            lbl���ƿ�Ŀ����.Enabled = True
            cmb���ƿ�Ŀ����.Enabled = True
        End If
    End If
    
    If mlng�������� > 0 And mlng�������� = Val(mstrID) Then
        If Item.Text = "��������" Or Item.Text = "��ҩ��" Then
            If Item.Checked = False Then
                MsgBox "�ò����ѱ�����ΪҽԺ����Һ�������ģ����ܸı����ԣ����ڻ������������д���", vbInformation, gstrSysName
                Item.Checked = True
                Exit Sub
            End If
        End If
    End If
    
    '���ѡ���ˡ�����ⷿ�����ԣ�����ѡ����������
    With lvw����
        For i = 1 To .ListItems.Count
            If .ListItems(i).Text = "����ⷿ" And .ListItems(i).Checked = True Then
                bln����ⷿ = True
                Exit For
            End If
        Next
        
        If bln����ⷿ = True Then
            For i = 1 To .ListItems.Count
                If .ListItems(i).Text <> "����ⷿ" And .ListItems(i).Checked = True Then
                    .ListItems(i).Checked = False
                End If
            Next
        End If
    End With
    
    Call Set�������Ҷ�Ӧ
End Sub

Private Sub lvw����_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lbl˵��.Caption = Item.ListSubItems(1).Tag
    mblnItem = True
End Sub

Private Sub lvw����_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("C") Or KeyAscii = Asc("c") Then Call ChangeServer
End Sub

Private Sub lvw����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lvw����.SelectedItem
        If .Tag = 0 Then
            lvw����.ToolTipText = "�ò��Ų������ڲ��ˣ��������ʲ����޸ģ�"
        Else
            lvw����.ToolTipText = "������ѡ��ʱ˫���򰴡�C�����ɸı�������"
        End If
    End With
End Sub

Private Sub lvw����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 2 Then
        If lvw����.SelectedItem Is Nothing Then Exit Sub
        If lvw����.SelectedItem.Checked = False Then Exit Sub
        If lvw����.SelectedItem.Tag = 0 Then Exit Sub
        
        For i = 0 To 4
            If mnuPatient(i).Caption <> "-" Then
                If lvw����.SelectedItem.SubItems(1) = Left(mnuPatient(i).Caption, InStr(mnuPatient(i).Caption, "(") - 1) Then
                    mnuPatient(i).Checked = True
                Else
                    mnuPatient(i).Checked = False
                End If
            End If
        Next
        PopupMenu mnuShort
    End If
End Sub

Private Sub mnuPatient_Click(Index As Integer)
    lvw����.SelectedItem.SubItems(1) = Left(mnuPatient(Index).Caption, InStr(mnuPatient(Index).Caption, "(") - 1)
    mblnChange = True
End Sub



Private Sub tabMain_Click()
    Dim i As Integer
    
    For i = fraMain.LBound To fraMain.UBound
        fraMain(i).Visible = False
    Next
    
    i = tabMain.SelectedItem.Index - 1
    fraMain(i).Visible = True
    fraMain(i).ZOrder 0
    If tabMain.SelectedItem.Index = 1 Then
        lblFind.Visible = False
        txtFind.Visible = False
    Else
        lblFind.Visible = True
        txtFind.Visible = True
    End If
End Sub
Private Sub ShowTab(ByVal intTab As Integer)
    tabMain.Tabs(intTab).Selected = True
    tabMain_Click
End Sub
Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(txtEdit(2).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = 2 Or Index = 5 Then
        OS.OpenIme True
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    ElseIf Index = 2 Or Index = 3 Then
        If LenB(StrConv(txtEdit(2).Text & Chr(KeyAscii), vbFromUnicode)) > 100 And (KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack) Then
            KeyAscii = 0
        End If
    ElseIf Index = 5 Then
        If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
    ElseIf Index = 0 Then
        If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Or Index = 5 Then
        OS.OpenIme False
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        If CheckOrder = True Then
            Cancel = True
        End If
    End If
End Sub

Private Function CheckOrder() As Boolean
    Dim rsTemp As Recordset
    Dim intOrder As Integer
    
    On Error GoTo ErrHandle
    
    If Val(txtEdit(0).Text) = 0 Then Exit Function
    CheckOrder = False
    
    If mstrID = "" Then '����
        If mstr�ϼ�����ID = "" Then
            gstrSQL = "Select 1 From ���ű� Where ˳�� = [1] And �ϼ�id is Null"
        Else
            gstrSQL = "Select 1 From ���ű� Where ˳�� = [1] And �ϼ�id =[2]"
        End If
    Else
        If mstr�ϼ�����ID = "" Then
            gstrSQL = "Select 1 From ���ű� Where ˳�� = [1] And �ϼ�id is Null And id <> [3]"
        Else
            gstrSQL = "Select 1 From ���ű� Where ˳�� = [1] And �ϼ�id =[2] And id <> [3]"
        End If
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����˳��", Val(txtEdit(0).Text), Val(mstr�ϼ�����ID), Val(mstrID))
    
    If Not rsTemp.EOF Then
        If mstr�ϼ�����ID = "" Then
            gstrSQL = "Select Max(Nvl(˳��,0)) As ���˳�� From ���ű� Where �ϼ�id Is Null"
        Else
            gstrSQL = "Select Max(Nvl(˳��,0)) As ���˳�� From ���ű� Where �ϼ�id = [1]"
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ���˳��", Val(mstr�ϼ�����ID))
        
        MsgBox "�ü���˳��Ϊ��" & Val(txtEdit(0).Text) & "���Ĳ����Ѵ��ڣ������˳��Ϊ��" & rsTemp!���˳�� & "��" & "�����������벿��˳��", vbInformation, gstrSysName
        CheckOrder = True
    End If
    
    rsTemp.Close
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    OS.OpenIme True
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte
    Dim strTemp As String
    Dim litem As ListItem
    Dim lsItem As ListItem
    Dim i As Integer
    
    If KeyAscii = 13 Then
        gstrSQL = " Select Distinct ����,����,ID From ���ű� " & _
         " Where ID in (Select ����ID From ��������˵�� Where " & mstr���� & ")" & _
         " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) and (���� like [1] or ���� like [2] or ���� like [3]) " & _
         " Order by ���� "
         
         strTemp = UCase(Trim(txtFind.Text))
         vRect = zlControl.GetControlRect(txtFind.hwnd)
         
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "���Ҳ�����Ӧ", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtFind.Height, True, False, True, strTemp & gstrLike, strTemp & gstrLike, strTemp & gstrLike)
                
        If Not rsTemp Is Nothing Then
            cmb���ƿ�Ŀ����.Text = rsTemp("����") & "-" & rsTemp("����")
            For Each litem In Lvw����.ListItems
                If litem.Text = cmb���ƿ�Ŀ����.Text Then
                    Set lsItem = litem
                Else
                    litem.Selected = False
                End If
            Next
            If Not lsItem Is Nothing Then
                lsItem.Selected = True
                txtFind.SetFocus
                txtFind.SelStart = 0
                txtFind.SelLength = Len(txtFind.Text)
                Exit Sub
            End If
        End If
        
        MsgBox "û���ҵ�����Ҫ�ÿ��ң������������ѯ������", vbInformation, gstrSysName
        txtFind.Text = ""
        txtFind.SetFocus
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(1).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(1).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

'CheckSameDept(txtEdit(2).Text, lngTmp) Then
Private Function CheckSameDept(ByVal strDept As String, ByRef lngDeptID As Long) As Boolean
'----------------------------------------------
'���ܣ������ɾ�������У��Ƿ�����ͬ�Ĳ�������
'������strDept����ǰ¼��Ĳ������ƣ�
'      lngDeptID��д���ҵ���ɾ��������ͬ�Ĳ���ID
'��ֵ��True���ҵ���ͬ��False��û���ҵ���
'----------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID From ���ű� " & _
              "Where Substr(����, 1, 1) = '-' And ����ʱ�� < To_Date('3000-1-1', 'yyyy-mm-dd') And ���� <> '��ɾ������' and ���� = [1] " & _
              "Order by ����ʱ�� desc "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ɾ�������뵱ǰ¼�벿���Ƿ���ͬ", strDept)
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp!ID) Then
            lngDeptID = rsTemp!ID
            CheckSameDept = True
        End If
    End If
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check�ظ�����(ByVal str�ϼ�ID As String, ByVal strҩƷ���� As String) As Boolean
    '���ܣ���������Ƿ��Ѿ��иò���
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select ���� from ���ű� where �ϼ�id=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�Ƿ����ظ�����", str�ϼ�ID, strҩƷ����)
    If rsTemp.EOF Then
        Check�ظ����� = False
    Else
        Check�ظ����� = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


