VERSION 5.00
Begin VB.Form frmSendArrange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ͱ���"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5040
   Icon            =   "frmSendArrange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancle 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "����(F2)"
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   5055
   End
   Begin VB.ComboBox cboRoom 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   2685
   End
   Begin VB.ComboBox cbo��ʦһ 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2685
   End
   Begin VB.ComboBox cbo��ʦ�� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   2685
   End
   Begin VB.Label lblRoom 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ  ��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   1750
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��鼼ʦ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   1075
      Width           =   1425
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��鼼ʦһ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   400
      Width           =   1425
   End
End
Attribute VB_Name = "frmSendArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngCurDeptId As Long
Private mlngAdviceId As Long
Private mlngSendNo As Long

Public Sub ShowMe(objParent As Object, ByVal lngCurDeptId As Long, ByVal lngAdviceId As Long, ByVal lngSendNo As Long)
    mlngCurDeptId = lngCurDeptId
    mlngAdviceId = lngAdviceId
    mlngSendNo = lngSendNo
    
    Me.Show 1, objParent
End Sub

Private Sub CmdCancle_Click()
    Unload Me
    
End Sub

Private Sub CmdOK_Click()
On Error GoTo ErrorHand

    Dim strSql As String
    
    strSql = "ZL_Ӱ�����¼_���Ͱ���(" & mlngAdviceId & "," & mlngSendNo & ",1," & "'" & NeedName(cbo��ʦһ.Text) & "','" & NeedName(cbo��ʦ��.Text) & "','" & NeedNo(cboRoom.Text) & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "���Ͱ���")
    
    '���汾�ε�ѡ��
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��鼼ʦһ", cbo��ʦһ.Text)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��鼼ʦ��", cbo��ʦ��.Text)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ִ�м�", cboRoom.Text)
    
    Unload Me
    
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHand
    Dim str��鼼ʦһ As String
    Dim str��鼼ʦ�� As String
    Dim strRoom As String

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    str��鼼ʦһ = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��鼼ʦһ")
    str��鼼ʦ�� = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��鼼ʦ��")
    strRoom = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ִ�м�")
    
    '���ؼ�鼼ʦ
    strSql = "Select " & vbNewLine & _
                "Distinct b.id,b.����, Upper(b.����) As ����" & vbNewLine & _
                " From ������Ա a, ��Ա�� b " & vbNewLine & _
                " Where a.��Աid = b.Id And " & vbNewLine & _
                "      (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) and a.����id = [1] " & vbNewLine & _
                " Order By ���� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    
    '���ؼ�鼼ʦһ
    cbo��ʦһ.Clear
    Do Until rsTmp.EOF
        cbo��ʦһ.AddItem rsTmp!���� & "-" & rsTmp!����
        
        If rsTmp!���� & "-" & rsTmp!���� = str��鼼ʦһ Then
            cbo��ʦһ.ListIndex = cbo��ʦһ.NewIndex
        End If
        
        If cbo��ʦһ.ListIndex = -1 And rsTmp!ID = UserInfo.ID Then
            cbo��ʦһ.ListIndex = cbo��ʦһ.NewIndex
        End If
        
        rsTmp.MoveNext
    Loop
    If cbo��ʦһ.ListCount > 0 And cbo��ʦһ.ListIndex = -1 Then cbo��ʦһ.ListIndex = 0
    
    '���ؼ�鼼ʦ��
    cbo��ʦ��.Clear
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Do Until rsTmp.EOF
            cbo��ʦ��.AddItem rsTmp!���� & "-" & rsTmp!����
            
            If rsTmp!���� & "-" & rsTmp!���� = str��鼼ʦ�� Then
                cbo��ʦ��.ListIndex = cbo��ʦ��.NewIndex
            End If
            
            If cbo��ʦ��.ListIndex = -1 And rsTmp!ID = UserInfo.ID Then
                cbo��ʦ��.ListIndex = cbo��ʦ��.NewIndex
            End If
            
            rsTmp.MoveNext
        Loop
        
        If cbo��ʦ��.ListCount > 0 And cbo��ʦ��.ListIndex = -1 Then cbo��ʦ��.ListIndex = 0
    End If
    
    'ִ�м�
    strSql = "Select ִ�м�,����豸 From ҽ��ִ�з��� Where ����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    
    cboRoom.Clear
    Do While Not rsTmp.EOF
        cboRoom.AddItem rsTmp!ִ�м� & "-" & Nvl(rsTmp!����豸)
        
        If Nvl(rsTmp!ִ�м�) & "-" & Nvl(rsTmp!����豸) = strRoom Then
            cboRoom.ListIndex = cboRoom.NewIndex
        End If
        
        rsTmp.MoveNext
    Loop
    If cboRoom.ListCount > 0 And cboRoom.ListIndex = -1 Then cboRoom.ListIndex = 0

    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
