VERSION 5.00
Begin VB.Form frmRegistPlan 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ת��"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "frmRegistPlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5385
   StartUpPosition =   1  '����������
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   5385
      TabIndex        =   9
      Top             =   0
      Width           =   5385
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   5500
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ָ��������Ҫת�ﵽ��Ŀ����ҵ���Ϣ��"
         Height          =   180
         Left            =   600
         TabIndex        =   11
         Top             =   390
         Width           =   3420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ת����Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   10
         Top             =   135
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4350
         Picture         =   "frmRegistPlan.frx":058A
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -150
      TabIndex        =   8
      Top             =   2700
      Width           =   6900
   End
   Begin VB.ComboBox cboҽ�� 
      Height          =   300
      Left            =   1905
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   2025
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   1905
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2010
      Width           =   2025
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      ItemData        =   "frmRegistPlan.frx":20CC
      Left            =   1905
      List            =   "frmRegistPlan.frx":20CE
      TabIndex        =   0
      Top             =   1125
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3660
      TabIndex        =   4
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2445
      TabIndex        =   3
      Top             =   2865
      Width           =   1100
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ת��ҽ��"
      Height          =   180
      Left            =   1125
      TabIndex        =   7
      Top             =   1620
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ת������"
      Height          =   180
      Left            =   1125
      TabIndex        =   6
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ת�����"
      Height          =   180
      Left            =   1125
      TabIndex        =   5
      Top             =   1185
      Width           =   720
   End
End
Attribute VB_Name = "frmRegistPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrNO As String

Private mlng����ID As Long
Private mstr���� As String
Private mstrҽ�� As String
Private mlngҽ��ID As Long

Private mstrԭ���� As String
Private mlngԭ����ID As Long

Private mlngPreDept As Long
Private mstrLike As String
Private mblnOK As Boolean

Public Function ShowMe(frmParent As Object, ByVal strNO As String, _
    lng����ID As Long, str���� As String, strҽ�� As String, lngҽ��ID As Long) As Boolean
'������strNO=Ҫת��ĹҺŵ�
'���أ�ת��ű�,ת�����,ת������,ת��ҽ����Ϣ
    
    mstrNO = strNO
    Me.Show 1, frmParent
    
    If mblnOK Then
        lng����ID = mlng����ID
        str���� = mstr����
        strҽ�� = mstrҽ��
        lngҽ��ID = mlngҽ��ID
    End If
    ShowMe = mblnOK
End Function

Private Sub cbo����_Click()
    If cbo����.ListIndex <> -1 Then
        If mlngPreDept <> cbo����.ItemData(cbo����.ListIndex) Then
            mlngPreDept = cbo����.ItemData(cbo����.ListIndex)
            '��ȡ�ÿ���ҽ��������
            Call LoadDoctor
            Call LoadRoom
        End If
    Else
        mlngPreDept = 0
    End If
End Sub

Private Sub cbo����_GotFocus()
    Call zlControl.TxtSelAll(cbo����)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
        
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If cbo����.Text <> "" Then
            strSql = "Select B.ID,B.����,B.����" & _
                " From ���ű� B,��������˵�� C" & _
                " Where B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
                " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
                " And (B.���� Like [1] Or Upper(B.����) Like [2] Or Upper(B.����) Like [2])" & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
                " Order by B.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(cbo����.Text) & "%", mstrLike & UCase(cbo����.Text) & "%")
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cbo����, rsTmp!ID)
            Else
                Call Cbo.SeekIndex(cbo����, mlngPreDept)
            End If
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboҽ��_Click()
    Call LoadRoom
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '�������
    If cbo����.ListIndex = -1 Then
        MsgBox "��ȷ��Ҫת��Ŀ��ҡ�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
'    If cbo����.Text = "" And cboҽ��.Text = "" Then
'        MsgBox "��ָ��ת�����һ�ҽ����", vbInformation, gstrSysName
'        If cbo����.Enabled Then cbo����.SetFocus
'        Exit Sub
'    End If
    If cbo����.ItemData(cbo����.ListIndex) = mlngԭ����ID Then
        If cboҽ��.Text = "" Then
            MsgBox "������ԭ������ת��ʱ����ָ��ת��ҽ����", vbInformation, gstrSysName
            If cboҽ��.Enabled Then cboҽ��.SetFocus
            Exit Sub
        End If
        If ZLCommFun.GetNeedName(cboҽ��.Text) = UserInfo.���� Then
            MsgBox "������ԭ������ת��ʱ��ת��ҽ��Ӧ��Ϊ����ҽ����", vbInformation, gstrSysName
            If cboҽ��.Enabled Then cboҽ��.SetFocus
            Exit Sub
        End If
    End If
    
    '��������
    mlng����ID = cbo����.ItemData(cbo����.ListIndex)
    mstr���� = cbo����.Text
    mstrҽ�� = ZLCommFun.GetNeedName(cboҽ��.Text)
    If cboҽ��.ListIndex <> -1 Then
        mlngҽ��ID = cboҽ��.ItemData(cboҽ��.ListIndex)
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Me.ActiveControl Is cbo���� Then
            KeyAscii = 0
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    mlng����ID = 0
    mstr���� = ""
    mstrҽ�� = ""
    mlngҽ��ID = 0
    mblnOK = False
    mlngPreDept = 0
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    
    On Error GoTo errH
    
    'ԭ�Һ������Ϣ
    strSql = "Select ִ�в���ID,���� From ���˹Һż�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrNO)
    mstrԭ���� = Nvl(rsTmp!����)
    mlngԭ����ID = rsTmp!ִ�в���ID
    
    '��ȡ�������:ȱʡΪ������
    strSql = "Select Distinct B.ID,B.����,B.����,Decode(B.ID,[1],1,0) as ȱʡ" & _
        " From ���ű� B,��������˵�� C" & _
        " Where B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
        " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
        " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngԭ����ID)
    Do While Not rsTmp.EOF
        cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
        If rsTmp!ȱʡ Then
            cbo����.ListIndex = cbo����.NewIndex '��������Click
            mlngPreDept = rsTmp!ID
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDoctor()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
                
    cboҽ��.Clear
    If cbo����.ListIndex = -1 Then Exit Sub
    
    strSql = "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
        " And C.��Ա����='ҽ��' And B.����ID=[1]" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, cbo����.ItemData(cbo����.ListIndex))
    
    cboҽ��.AddItem ""
    Call Cbo.SetIndex(cboҽ��.hwnd, 0)
    Do While Not rsTmp.EOF
        cboҽ��.AddItem rsTmp!���� & "-" & rsTmp!����
        cboҽ��.ItemData(cboҽ��.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRoom()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim bytRegistMode As Byte, datRegistTime As Date
    
    On Error GoTo errH
    
    cbo����.Clear
    If cbo����.ListIndex = -1 Then Exit Sub
    
    bytRegistMode = Val(Split(zlDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(0))
    If Split(zlDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(1) <> "" Then
        datRegistTime = CDate(Format(Split(zlDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(1), "yyyy-mm-dd hh:mm:ss"))
    End If
    
    If bytRegistMode = 0 Then
        strSql = "Select ID From �ҺŰ��� Where ����ID=[1] And (ҽ������=[2] Or ҽ������ Is Null Or [2] Is Null)"
        strSql = "Select Distinct �������� From �ҺŰ������� Where �ű�ID IN(" & strSql & ") Order by ��������"
    Else
        If Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") < Format(datRegistTime, "yyyy-mm-dd hh:mm:ss") Then
            strSql = "Select ID From �ҺŰ��� Where ����ID=[1] And (ҽ������=[2] Or ҽ������ Is Null Or [2] Is Null)"
            strSql = "Select Distinct �������� From �ҺŰ������� Where �ű�ID IN(" & strSql & ") Order by ��������"
        Else
            strSql = "Select A.ID From �ٴ������¼ A Where A.����ID=[1] And (A.ҽ������=[2] Or A.ҽ������ Is Null Or [2] Is Null)"
            strSql = "Select Distinct B.���� As �������� From �ٴ��������Ҽ�¼ A,�������� B Where A.����ID=B.ID And A.��¼ID IN(" & strSql & ") Order by B.����"
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, cbo����.ItemData(cbo����.ListIndex), ZLCommFun.GetNeedName(cboҽ��.Text))
    
    cbo����.AddItem ""
    Call Cbo.SetIndex(cbo����.hwnd, 0)
    Do While Not rsTmp.EOF
        cbo����.AddItem rsTmp!��������
        If cbo����.ItemData(cbo����.ListIndex) = mlngԭ����ID And rsTmp!�������� = mstrԭ���� Then
            cbo����.ListIndex = cbo����.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
