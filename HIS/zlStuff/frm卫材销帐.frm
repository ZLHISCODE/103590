VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frm�������� 
   Caption         =   "������������"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11760
   Icon            =   "frm��������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   11760
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   90
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   19
      Top             =   2295
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   15
         TabIndex        =   20
         Top             =   30
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "��������(&V)"
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
      Left            =   12135
      TabIndex        =   11
      ToolTipText     =   "�ȼ���F2"
      Top             =   90
      Width           =   1335
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
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
      Left            =   12135
      TabIndex        =   10
      Top             =   825
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
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
      Left            =   12135
      TabIndex        =   9
      ToolTipText     =   "�ȼ���F2"
      Top             =   465
      Width           =   1335
   End
   Begin VB.Frame fraCondition 
      Height          =   1125
      Left            =   30
      TabIndex        =   1
      Top             =   75
      Width           =   11985
      Begin VB.TextBox txtPati 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9540
         TabIndex        =   17
         ToolTipText     =   "����סԺ�š�����ID������(ָ���˲���ʱ)"
         Top             =   165
         Width           =   2355
      End
      Begin VB.ComboBox cbo������ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8250
         TabIndex        =   15
         Text            =   "cbo������"
         Top             =   645
         Width           =   2310
      End
      Begin VB.ComboBox cbo���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4425
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   645
         Width           =   2790
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ҽ������(&W)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2670
         TabIndex        =   7
         Top             =   690
         Width           =   1575
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "����(&T)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1470
         TabIndex        =   6
         Top             =   690
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker Dtp��ʼʱ�� 
         Height          =   315
         Left            =   1350
         TabIndex        =   2
         Top             =   195
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   121831427
         CurrentDate     =   36985
      End
      Begin MSComCtl2.DTPicker Dtp����ʱ�� 
         Height          =   315
         Left            =   4455
         TabIndex        =   3
         Top             =   195
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   121831427
         CurrentDate     =   36985
      End
      Begin VB.CheckBox chk�����ڼ� 
         Caption         =   "�����ڼ�(&S)"
         Height          =   195
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   1290
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10710
         TabIndex        =   8
         ToolTipText     =   "�ȼ���F2"
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "�����ڼ�(&S)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblPatiInputType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ�š�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   8685
         TabIndex        =   18
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7725
         TabIndex        =   16
         Top             =   210
         Width           =   840
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7530
         TabIndex        =   14
         Top             =   705
         Width           =   630
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   465
         TabIndex        =   13
         Top             =   705
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4200
         TabIndex        =   5
         Top             =   255
         Width           =   210
      End
   End
   Begin VB.Menu mnuPati 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuPatiItem 
         Caption         =   "סԺ��(&0)"
         Index           =   0
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "ID(&1)"
         Index           =   1
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "����(&2)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�ӿڲ���
Private mintUnit As String              '0-ɢװ��λ,1-��װ��λ
Private mint����λ�� As Integer
Private mlng���ϲ���ID As Long
Private mArrFilter As Variant   '��������
Private mstrPrivs As String
Private mlngModule As Long
'��������
Private mblnDrop As Boolean                     '��KeyDown���ж������б��Ƿ񵯳�
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private mfrmδ�� As frm��������_δ���
Private mfrm���� As frm��������_�����
Private mstr��ʼ����ʱ�� As String, mstr��������ʱ�� As String
Private mstr��ʼ���ʱ�� As String, mstr�������ʱ�� As String

Private Enum mPage
    pag_δ�� = 0
    pag_���� = 1
End Enum

Private mobjPlugIn As Object             '��ҽӿڶ���

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Private Function GetFilter() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-30 11:52:50
    '-----------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, strReg As String
       
    '������ѯ����
    Set cllFilter = New Collection
    If mlng���ϲ���ID < 0 Then
        cllFilter.Add 0, "���ϲ���ID"
    Else
        cllFilter.Add mlng���ϲ���ID, "���ϲ���ID"
    End If
    If cbo����.ListIndex <= 0 Then
        cllFilter.Add 0, "�������ID"
    Else
        cllFilter.Add cbo����.ItemData(cbo����.ListIndex), "�������ID"
    End If
    
    cllFilter.Add Array("1949-01-01 00:00:00", "1949-01-01 23:59:59"), "���ڷ�Χ"
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_δ��
        If chk�����ڼ�.Value = 1 Then
            cllFilter.Remove "���ڷ�Χ"
            cllFilter.Add Array(Format(dtp��ʼʱ��.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtp����ʱ��.Value, "yyyy-mm-dd HH:MM:SS")), "���ڷ�Χ"
            mstr��ʼ����ʱ�� = Format(dtp��ʼʱ��.Value, "yyyy-mm-dd HH:MM:SS")
            mstr��������ʱ�� = Format(dtp����ʱ��.Value, "yyyy-mm-dd HH:MM:SS")
        End If
    Case mPage.pag_����
        cllFilter.Add Array(Format(dtp��ʼʱ��.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtp����ʱ��.Value, "yyyy-mm-dd HH:MM:SS")), "�������"
        mstr��ʼ���ʱ�� = Format(dtp��ʼʱ��.Value, "yyyy-mm-dd HH:MM:SS")
        mstr�������ʱ�� = Format(dtp����ʱ��.Value, "yyyy-mm-dd HH:MM:SS")
    End Select
    
    If cbo������.ListIndex = 0 Then
        cllFilter.Add "", "������"
    Else
        cllFilter.Add NeedName(cbo������.Text), "������"
    End If
    cllFilter.Add "", "��������"
    
    If Trim(txtPati.Text) <> "" Then
        If Val(lblPatiInputType.Tag) = 0 Then
            cllFilter.Add Val(txtPati.Tag), "סԺ��"
        Else
            cllFilter.Add 0, "סԺ��"
        End If
        
        If Val(lblPatiInputType.Tag) = 1 Then
            cllFilter.Add Val(txtPati.Tag), "����ID"
        Else
            cllFilter.Add 0, "����ID"
        End If
        If Val(lblPatiInputType.Tag) = 2 Then
            cllFilter.Add Trim(txtPati.Tag), "����"
        Else
            cllFilter.Add "", "����"
        End If
    Else
        cllFilter.Add 0, "סԺ��"
        cllFilter.Add 0, "����ID"
        cllFilter.Add "", "����"
    End If
 
   ' Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", _
          Val(mArrFilter("���ϲ���id")), Val(mArrFilter("�������id")), _
          CDate(mArrFilter("ʱ�䷶Χ")(0)), CDate(mArrFilter("ʱ�䷶Χ")(1)), _
          CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1)), _
          Trim(mArrFilter("������")), Trim(mArrFilter("��������")), _
          Val(mArrFilter("סԺ��")), Val(mArrFilter("����ID")))
        
    
    Set mArrFilter = cllFilter
    
End Function

Private Sub ��ȡ������()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��ز��ŵ�������
    '���:int��������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-03 23:25:54
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    Dim strWhere As String
    
    On Error GoTo ErrHandle
    If cbo����.ListIndex > 0 Then strWhere = " And B.����id = [1] "
 
    gstrSQL = "" & _
        "   Select Distinct A.ID, A.����||'-'||A.���� As ���� " & _
        "   From ��Ա�� A, ������Ա B " & _
        "   Where A.ID = B.��Աid And (a.վ��=[2] or a.վ�� is null) " & strWhere & _
        "           And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
        "   Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ա", Val(cbo����.ItemData(cbo����.ListIndex)), gstrNodeNo)
        
    cbo������.Clear
    cbo������.AddItem "����������"
    cbo������.ItemData(cbo������.NewIndex) = 0
    Do While Not rsTemp.EOF
        cbo������.AddItem rsTemp!����
        cbo������.ItemData(cbo������.NewIndex) = rsTemp!Id
        rsTemp.MoveNext
    Loop
    cbo������.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub ��ȡ���ϲ�������()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ���ϲ�������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-05-03 23:29:52
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ���� From ���ű� Where ID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ⷿ����", mlng���ϲ���ID)
    
    If Not rsTemp.EOF Then
        Me.Caption = Me.Caption & "(��ǰ�ⷿ��" & rsTemp!���� & ")"
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowList(frmMain As Form, strPirvs As String, lngModule As Long, ByVal lng���ϲ���ID As Long, ByVal intUnit As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������
    '���:lng���ϲ���ID-���ϲ���ID
    '     intUnit-��ʾ��λ(0-ɢװ��λ,1-��װ��λ)
    '     int����λ��-����λ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-03 23:46:24
    '-----------------------------------------------------------------------------------------------------------
    mstrPrivs = strPirvs: mlngModule = lngModule
    mlng���ϲ���ID = lng���ϲ���ID
    mintUnit = intUnit
    'mint����λ�� = int����λ��
    Me.Show vbModal, frmMain
    ShowList = True
End Function

Private Sub cbo����_Click()
    If cbo����.ListIndex = -1 Then Exit Sub
    If Val(cbo����.Tag) <> cbo����.ItemData(cbo����.ListIndex) Then
        cbo����.Tag = cbo����.ItemData(cbo����.ListIndex)
        Call ��ȡ������
    End If
End Sub

Private Sub cbo������_Click()
    Exit Sub
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String, strResult As String, strFilter As String

    If KeyAscii = 13 Then
        strText = UCase(cbo������.Text)
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo������.List(cbo������.ListIndex) Then Call zlControl.CboSetIndex(cbo������.hwnd, -1)
        End If
        If strText = "" Then
            cbo������.ListIndex = -1
        ElseIf cbo������.ListIndex = -1 Then
            intIdx = -1

            For i = 1 To cbo������.ListCount - 1
                If Mid(cbo������.List(i), 1, InStr(1, cbo������.List(i), "-") - 1) = strText _
                    Or Mid(cbo������.List(i), InStr(1, cbo������.List(i), "-")) = strText Then
                    intIdx = i
                    Exit For
                End If
            Next

            If intIdx = -1 Then
                For i = 1 To cbo������.ListCount - 1
                    If UCase(cbo������.List(i)) Like strText & "*" Then
                        intIdx = i
                    End If
                Next
            End If

            cbo������.ListIndex = intIdx
            SendMessage cbo������.hwnd, CB_SHOWDROPDOWN, True, 0
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo������_Click
            Exit Sub
        End If
        If cbo������.ListIndex = -1 Then
            cbo������.ListIndex = 0
        Else
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo������_Click
            ElseIf intIdx <> cbo������.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo������.SetFocus
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo������_Click
            End If
        End If
    End If
End Sub
Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Private Sub chk�����ڼ�_Click()
    dtp��ʼʱ��.Enabled = chk�����ڼ�.Value = 1
    dtp����ʱ��.Enabled = chk�����ڼ�.Value = 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub
Private Sub IniDate()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ��������ڵ�Ĭ��ֵ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-05-03 23:51:57
    '-----------------------------------------------------------------------------------------------------------
    Dim dtCurrDate As Date
    dtCurrDate = sys.Currentdate
    dtp��ʼʱ��.MaxDate = CDate(Format(dtCurrDate, "yyyy-MM-dd 23:59:59"))
    dtp����ʱ��.MaxDate = dtp��ʼʱ��.MaxDate
    dtp��ʼʱ��.Value = CDate(Format(DateAdd("D", -1, dtCurrDate), "yyyy-MM-dd 00:00:00"))
    dtp����ʱ��.Value = CDate(Format(dtCurrDate, "yyyy-MM-dd 23:59:59"))
    mstr��ʼ����ʱ�� = Format(dtp��ʼʱ��.Value, "yyyy-mm-dd HH:MM:SS")
    mstr��������ʱ�� = Format(dtp����ʱ��.Value, "yyyy-mm-dd HH:MM:SS")
    mstr��ʼ���ʱ�� = mstr��ʼ����ʱ��
    mstr�������ʱ�� = mstr��������ʱ��
    
End Sub
Private Sub ��ȡ��������(ByVal int�������� As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-05-03 23:52:53
    '-----------------------------------------------------------------------------------------------------------

    'int�������ͣ�0-������1-ҽ������
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    Select Case int��������
        Case 0
            gstrSQL = " Select ����||'-'||���� ����,ID From ���ű� " & _
             " Where ID in (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3))" & _
             "     And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) And (վ��=[1] or վ�� is null) " & _
             " Order By ����||'-'||���� "
        Case 1
            gstrSQL = " Select ����||'-'||���� ����,ID From ���ű� " & _
             " Where ID in (Select ����ID From ��������˵�� Where �������� In ('���','����','����','����') And ������� IN(2,3))" & _
             "     And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) And (վ��=[1] or վ�� is null) " & _
             " Order By ����||'-'||���� "
    End Select
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", gstrNodeNo)
    
    cbo����.Clear
    
    If int�������� = 0 Then
        cbo����.AddItem "���в���"
        cbo����.ItemData(cbo����.NewIndex) = 0
    Else
        cbo����.AddItem "���п���"
        cbo����.ItemData(cbo����.NewIndex) = 0
    End If
    
    Do While Not rsTemp.EOF
        cbo����.AddItem rsTemp!����
        cbo����.ItemData(cbo����.NewIndex) = rsTemp!Id
        rsTemp.MoveNext
    Loop
    
    cbo����.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub IniDept()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-05-03 23:53:39
    '-----------------------------------------------------------------------------------------------------------
    If Lbl����.Tag = "" Then
        Lbl����.Tag = "-1"
        opt����_Click (0)
    End If
End Sub
Private Function FullData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-03 23:57:13
    '-----------------------------------------------------------------------------------------------------------
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_δ��
        FullData = mfrmδ��.zlRefreshData(Me, mstrPrivs, mlngModule, mintUnit, mArrFilter)
    Case mPage.pag_����
        FullData = mfrm����.zlRefreshData(Me, mstrPrivs, mlngModule, mintUnit, mArrFilter)
    End Select
End Function
Private Sub cmdRefresh_Click()
    Call GetFilter
    Call FullData
End Sub

Private Sub cmdVerify_Click()
    Set mfrmδ��.In_PlugIn = mobjPlugIn
    If mfrmδ��.zlVerifyData = False Then Exit Sub
    Call cmdRefresh_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    cbo����.Tag = "-1"
    Call IniDate
    Call IniDept
    Call ��ȡ���ϲ�������
    Call InitPage
'    Call GetFilter
End Sub

Private Sub Form_Resize()
    Dim lngTmp As Long
    
    If WindowState = 1 Then Exit Sub
    On Error Resume Next
    If Me.Width < 13860 Then Me.Width = 13860
    If Me.Height < 8790 Then Me.Height = 8790
    
    cmdVerify.Left = Me.ScaleWidth - cmdVerify.Width - 50
    cmdExit.Left = cmdVerify.Left
    cmdHelp.Left = cmdVerify.Left
    fraCondition.Left = Me.ScaleLeft + 20
    fraCondition.Width = Me.ScaleWidth - cmdVerify.Width - 100
    With picList
        .Top = fraCondition.Height + fraCondition.Top + 50
        .Height = Me.ScaleHeight - .Top - 50
        .Width = Me.ScaleWidth - .Left - 50
    End With
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmδ�� Is Nothing Then
        Unload mfrmδ��
        Set mfrmδ�� = Nothing
    End If
    
    If Not mfrm���� Is Nothing Then
        Unload mfrm����
        Set mfrm���� = Nothing
    End If
End Sub

Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        PopupMenu mnuPati, 2, lblPatiInputType.Left + lblPatiInputType.Width - 30, lblPatiInputType.Top
    End If
End Sub

Private Sub mnuPatiItem_Click(Index As Integer)
    Select Case Index
        Case 0
            lblPatiInputType.Caption = "סԺ�š�"
            lblPatiInputType.Tag = 0
            txtPati.Text = ""
            txtPati.Tag = ""
        Case 1
            lblPatiInputType.Caption = "ID��"
            lblPatiInputType.Tag = 1
            txtPati.Text = ""
            txtPati.Tag = ""
        Case 2
            lblPatiInputType.Caption = "���š�"
            lblPatiInputType.Tag = 2
            txtPati.Text = ""
            txtPati.Tag = ""
    End Select
End Sub
Private Sub opt����_Click(Index As Integer)
    If Val(Lbl����.Tag) <> Index Then
        If Index = 1 Then
            mnuPatiItem(2).Enabled = False
            If Val(lblPatiInputType.Tag) = 2 Then
                Call mnuPatiItem_Click(0)
            End If
        Else
            mnuPatiItem(2).Enabled = True
        End If
        
        Call ��ȡ��������(Index)
        Lbl����.Tag = Index
    End If
End Sub

Private Sub picList_Resize()
    With tbPage
        .Top = picList.ScaleTop
        .Height = picList.ScaleHeight
        .Width = picList.ScaleWidth
        .Left = picList.ScaleLeft
    End With

End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call SetInitDate(Val(Item.Tag))
End Sub
Private Sub SetInitDate(ByVal intType As Integer)
    '����:����ʱ����ʾ
    Select Case intType
    Case mPage.pag_����
        cmdVerify.Enabled = False
         lblʱ��.Caption = "����ڼ�(&S)"
        chk�����ڼ�.Visible = False
        lblʱ��.Visible = True
        dtp��ʼʱ��.Value = CDate(mstr��ʼ���ʱ��)
        dtp����ʱ��.Value = CDate(mstr�������ʱ��)
        dtp��ʼʱ��.Enabled = True
        dtp����ʱ��.Enabled = True
    Case Else
        lblʱ��.Caption = "�����ڼ�(&S)"
        cmdVerify.Enabled = True
        chk�����ڼ�.Visible = True
        dtp��ʼʱ��.Enabled = chk�����ڼ�.Value = 1: dtp����ʱ��.Enabled = chk�����ڼ�.Value = 1
        lblʱ��.Visible = False
        dtp��ʼʱ��.Value = CDate(mstr��ʼ����ʱ��)
        dtp����ʱ��.Value = CDate(mstr��������ʱ��)
        dtp��ʼʱ��.Enabled = chk�����ڼ�.Value = 1
        dtp����ʱ��.Enabled = chk�����ڼ�.Value = 1
    End Select
End Sub

Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim str��ʶ�� As String
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    txtPati.Text = Trim(txtPati.Text)
    
    If txtPati.Text = "" Then Exit Sub
    
    If Val(lblPatiInputType.Tag) = 0 Then
        If InStr(1, txtPati.Text, "-") > 0 Then
            str��ʶ�� = Mid(txtPati.Text, 1, InStr(1, txtPati.Text, "-") - 1)
        Else
            str��ʶ�� = txtPati.Text
        End If
        gstrSQL = "Select Distinct ����,סԺ�� As ��ʶ From ������Ϣ Where סԺ�� = [1] "
    ElseIf Val(lblPatiInputType.Tag) = 1 Then
        gstrSQL = "Select Distinct ����,����ID As ��ʶ From ������Ϣ Where ����ID = [2] "
    Else
        If cbo����.ListIndex = 0 Then
            MsgBox "��ѡ������"
            Exit Sub
        End If
        str��ʶ�� = txtPati.Text
        gstrSQL = "Select A.����,B.���� As ��ʶ From ������Ϣ A, ��λ״����¼ B Where A.����id = B.����id And ����id = [3] And B.���� = [1] "
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", str��ʶ��, Val(txtPati.Text), Val(cbo����.ItemData(cbo����.ListIndex)))
    If rsTemp.RecordCount > 0 Then
        txtPati.Text = rsTemp!��ʶ & "-" & rsTemp!����
        txtPati.Tag = rsTemp!��ʶ
        
        cmdRefresh_Click
    Else
        'ͨ������һ�������ڵ���������ձ������
        txtPati.Text = "-1"
        txtPati.Tag = "-1"
        cmdRefresh_Click
        
        txtPati.Text = ""
        txtPati.Tag = ""
        txtPati.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If Val(lblPatiInputType.Tag) = 0 Or Val(lblPatiInputType.Tag) = 1 Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Then Exit Sub
        KeyAscii = 0
    End If
End Sub
   
Private Sub InitPage()
    '------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:
    '����:���˺�
    '����:2007/08/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim objItem As TabControlItem
 
    
    Set mfrmδ�� = New frm��������_δ���
    Set objItem = tbPage.InsertItem(mPage.pag_δ��, "δ���", mfrmδ��.hwnd, 0)
    objItem.Tag = mPage.pag_δ��
    Set mfrm���� = New frm��������_�����
    Set objItem = tbPage.InsertItem(mPage.pag_����, "�����", mfrm����.hwnd, 0)
    objItem.Tag = mPage.pag_����
    With tbPage
        .Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
    End With
    Call SetInitDate(mPage.pag_δ��)
    Call GetFilter
    Call FullData
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

