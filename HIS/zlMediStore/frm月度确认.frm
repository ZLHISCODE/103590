VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm�¶�ȷ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�¶�ȷ��"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frm�¶�ȷ��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd��� 
      Caption         =   "���(&L)"
      Height          =   350
      Left            =   150
      TabIndex        =   9
      ToolTipText     =   "�����һ���µ�ȷ�ϼ�¼"
      Top             =   2010
      Width           =   945
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3210
      TabIndex        =   8
      Top             =   2010
      Width           =   945
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2190
      TabIndex        =   7
      Top             =   2010
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "ȷ�ϱ��¶ȵ���ֹ����(&A)"
      Height          =   1755
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4155
      Begin VB.ComboBox cbo��ǰ�·� 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtp��ʼ���� 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   780
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   114491395
         CurrentDate     =   38148
      End
      Begin MSComCtl2.DTPicker dtp�������� 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   1230
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   114491395
         CurrentDate     =   38148
      End
      Begin VB.Label lbl�·� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�·�(&M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   1
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   1297
         Width           =   990
      End
      Begin VB.Label lbl��ʼ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   847
         Width           =   990
      End
   End
End
Attribute VB_Name = "frm�¶�ȷ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrȱʡ�������� As String
Private mlng�ⷿID As Long
Private mblnStart As Boolean

Private Sub cbo��ǰ�·�_Click()
    Dim str��ʼ���� As String, str�������� As String
    
    Call GetPeriod(Me.cbo��ǰ�·�.Text, str��ʼ����, str��������)
    '���ý������ںͿ�ʼ����
    '�������ڵ���Сʱ��Ϊ��ʼʱ���1��
    Me.Dtp��������.Value = Format(str��������, "yyyy��MM��dd�� HH:mm:ss")
    Me.Dtp��ʼ����.Value = Format(str��ʼ����, "yyyy��MM��dd�� HH:mm:ss")
    Me.Dtp��ʼ����.Enabled = (Me.cbo��ǰ�·�.ListCount > 2)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    On Error GoTo ErrHand
    
    '��ʼ���ڱ���С�ڽ�������
    If Not (Format(Me.Dtp��������.Value, "yyyy-MM-dd HH:mm:ss") > Format(Me.Dtp��ʼ����.Value, "yyyy-MM-dd HH:mm:ss")) Then
        MsgBox "�������ڱ�����ڿ�ʼʱ�䣡", vbInformation, gstrSysName
        Me.Dtp��������.SetFocus
        Exit Sub
    End If
    If Format(Me.Dtp��������.Value, "yyyy-MM-dd HH:mm:ss") > Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss") Then
        MsgBox "�������ڲ��ܴ��ڵ�ǰ���ڣ�", vbInformation, gstrSysName
        Me.Dtp��������.SetFocus
        Exit Sub
    End If
    
    gstrSQL = "zl_�ⷿȷ�ϼ�¼_UPDATE(" & mlng�ⷿID & ",'" & Me.cbo��ǰ�·�.Text & "'" & _
        ",to_date('" & Format(Me.Dtp��ʼ����.Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')" & _
        ",to_date('" & Format(Me.Dtp��������.Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')" & _
        ")"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "����ⷿȷ�ϼ�¼")
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd���_Click()
    On Error GoTo ErrHand
    
    If MsgBox("��ȷ��Ҫ������һ���µ�ȷ�ϼ�¼��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "zl_�ⷿȷ�ϼ�¼_Back(" & mlng�ⷿID & ")"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "������һ�ε�ȷ�ϼ�¼")
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim str�·� As String
    Dim blnInit As Boolean
    Dim rsTemp As New ADODB.Recordset
    '��һ�ο����ɲ���Ա���ÿ�ʼ���ڡ�ѡ���·ݣ�ȱʡ�Ŀ�ʼ���ںͽ�������Ϊ�ڼ��Ŀ�ʼ���ںͽ�������
    '�Ժ�ֻ������
    '   1���������ڣ���������ȱʡΪ���죬���ѡ�񵽵��죬������ǰѡ�񣬵�����С�ڵ��ڿ�ʼ����
    '   2���·ݲ���ѡ��ֻ���Ǵ������һ���·�
    On Error GoTo errHandle
    mblnStart = False
    'ȱʡװ�����¡����������·�
    gstrSQL = "" & _
        " Select MAX(�·�) �·�" & _
        " From �ⷿȷ�ϼ�¼" & _
        " WHERE �ⷿID=[1] And ����=1"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ����·�]", mlng�ⷿID)
        
    If Not IsNull(rsTemp!�·�) Then
        blnInit = False
        str�·� = rsTemp!�·�   '����·�=����
    Else
        blnInit = True
    End If
    
    'װ���ڼ����һ����װ�������ڼ䣬����װ�����¼��Ժ��ڼ䣩
    gstrSQL = "Select �ڼ� As �·�,��ʼ����,��ֹ���� " & _
            " From �ڼ��" & _
            " Where " & IIf(blnInit, "1=1", " �ڼ�>=[1] And Rownum<3") & _
            "" & _
            " Order by �ڼ�"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ڼ��]", str�·�)
    
    Me.cbo��ǰ�·�.Clear
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
    Else
        Exit Sub
    End If
    
    Do While Not rsTemp.EOF
        Me.cbo��ǰ�·�.AddItem rsTemp!�·�
        rsTemp.MoveNext
    Loop
    
    '��ȡ�ϴν�������
    Call LocateCbo(str�·�)
    
    mblnStart = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowEditor(ByVal lng�ⷿID As Long, Optional ByVal str�������� As String = "")
    mlng�ⷿID = lng�ⷿID
    mstrȱʡ�������� = str��������
    Me.Show 1
End Sub

Private Function LocateCbo(ByVal strInput As String) As Boolean
    Dim intItem As Integer, intItems As Integer
    '��λ����ǰ�·�
    LocateCbo = True
    Me.cbo��ǰ�·�.ListIndex = 0
    If strInput = "" Then Exit Function
    intItems = Me.cbo��ǰ�·�.ListCount - 1
    For intItem = 0 To intItems
        If Me.cbo��ǰ�·�.Text = strInput Then
            Me.cbo��ǰ�·�.ListIndex = intItem
            Exit Function
        End If
    Next
    LocateCbo = False
End Function

Private Sub GetPeriod(ByVal str�·� As String, str��ʼ���� As String, str�������� As String)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    '��ȡָ���·ݵĿ�ʼ���ں���ֹ����
    gstrSQL = "Select to_char(��ʼʱ��,'yyyy-MM-DD hh24:mi:ss') As ��ʼ����,to_char(��ֹʱ��,'yyyy-MM-DD hh24:mi:ss') As ��������" & _
            " From �ⷿȷ�ϼ�¼" & _
            " Where �ⷿID=[1] And �·�=[2] And ����=1"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡָ���·ݵĿ�ʼ���ں���ֹ����]", mlng�ⷿID, str�·�)
    
    If Not rsTemp.EOF Then
        str��ʼ���� = rsTemp!��ʼ����
        str�������� = rsTemp!��������
    Else
        '����ȥȡ�ϸ��µ���ֹ������Ϊ���ο�ʼ���ڣ��������ȡ������˵���ǵ�һ������
        Call GetStartDate(str�·�, str��ʼ����)
        If mstrȱʡ�������� <> "" And mstrȱʡ�������� > str��ʼ���� Then
            str�������� = mstrȱʡ��������
        Else
            str�������� = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetStartDate(ByVal str�·� As String, str��ʼ���� As String)
    Dim rsTemp As New ADODB.Recordset
    '�����ϸ��µ���ֹ���ڵõ����¿�ʼ����
    On Error GoTo errHandle
    gstrSQL = "Select to_char(max(��ֹʱ��),'yyyy-MM-DD hh24:mi:ss') As ��ʼ����" & _
            " From �ⷿȷ�ϼ�¼" & _
            " Where �ⷿID=[1] And �·�<[2] And ����=1"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[�����ϸ��µ���ֹ���ڵõ����¿�ʼ����]", mlng�ⷿID, str�·�)
    
    If Not IsNull(rsTemp!��ʼ����) Then
        str��ʼ���� = DateAdd("s", 1, rsTemp!��ʼ����)
    Else
        str��ʼ���� = "2004-01-01 00:00:00"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
