VERSION 5.00
Begin VB.Form frmRunLimitPlanEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2745
      TabIndex        =   4
      Top             =   1875
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1485
      TabIndex        =   2
      Top             =   1875
      Width           =   1100
   End
   Begin VB.Frame fraMain 
      Height          =   1755
      Left            =   120
      TabIndex        =   3
      Top             =   15
      Width           =   3750
      Begin VB.TextBox txtPlanName 
         Height          =   300
         Left            =   900
         TabIndex        =   0
         Top             =   225
         Width           =   2730
      End
      Begin VB.TextBox txtPlanDescription 
         Height          =   990
         Left            =   900
         MaxLength       =   125
         TabIndex        =   1
         Top             =   660
         Width           =   2730
      End
      Begin VB.Label lblPlanName 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   60
         TabIndex        =   6
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblPlanDescription 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   60
         TabIndex        =   5
         Top             =   660
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmRunLimitPlanEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngPlanNo As Long
Private mstrPlanName As String
Private mstrDescription As String
Private mblnOk As Boolean

Public Function ShowMe(ByVal frmFather As Object, ByRef lngPlanNo As Long, ByRef strPlanName As String, ByRef strDescription As String) As Boolean
    mlngPlanNo = lngPlanNo
    mstrPlanName = strPlanName
    mstrDescription = strDescription
    Me.Show vbModal, frmFather
    If mblnOk Then
        lngPlanNo = mlngPlanNo
        strPlanName = mstrPlanName
        strDescription = mstrDescription
        ShowMe = mblnOk
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errH
    mblnOk = False
    txtPlanName.Text = Trim(txtPlanName.Text)
    If CheckData Then
        If mlngPlanNo = 0 Then
            '����
            Call ExecuteProcedure("Zl_Zlrunlimit_Update(0,Null,'" & txtPlanName.Text & "',1,'" & txtPlanDescription.Text & "')", "��������")
            strSql = "Select ��� From Zlrunlimit Where ���� = [1]"
            Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "��ȡ�ղ������ķ��������", txtPlanName.Text)
            mlngPlanNo = rsTemp!���
        Else
            '�޸�
            Call ExecuteProcedure("Zl_Zlrunlimit_Update(1," & mlngPlanNo & ",'" & txtPlanName.Text & "',Null,'" & txtPlanDescription.Text & "')", "�޸ķ���")
        End If
        mstrPlanName = txtPlanName.Text
        mstrDescription = txtPlanDescription.Text
        mblnOk = True
        Unload Me
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'����������ݵĺϷ���
Private Function CheckData() As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If txtPlanName.Text = "" Then
        MsgBox "�������Ʋ���Ϊ�գ���������д��", vbInformation, gstrSysName
        txtPlanName.SetFocus
        Exit Function
    End If
    If txtPlanName.Text = "[�޷�������]" Then
        MsgBox "������Ϊϵͳ�������ƣ�����ʹ�ã���������д��", vbInformation, gstrSysName
        txtPlanName.SetFocus
        Exit Function
    End If
    If InStr(txtPlanName.Text, "'") > 0 Then
        MsgBox "�����в��ܺ��е�����,��������д��", vbInformation, gstrSysName
        txtPlanName.SetFocus
        Exit Function
    End If
    If InStr(txtPlanDescription.Text, "'") > 0 Then
        MsgBox "�����в��ܺ��е�����,��������д��", vbInformation, gstrSysName
        txtPlanDescription.SetFocus
        Exit Function
    End If
    If StrIsValid(txtPlanName.Text, 50) = False Then
        txtPlanName.SetFocus
        Exit Function
    End If
    If StrIsValid(txtPlanDescription.Text, 250) = False Then
        txtPlanDescription.SetFocus
        Exit Function
    End If
    If mstrPlanName <> txtPlanName.Text Then
        strSql = "Select Count(1) ���� From Zlrunlimit Where ���� = [1]"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption, txtPlanName.Text)
        If rsTemp!���� = 1 Then
            MsgBox "�÷��������Ѿ����ڣ���������д��", vbInformation, gstrSysName
            txtPlanName.SetFocus
            Exit Function
        End If
    End If
    CheckData = True
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    '��ֹ���뵥����
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    If mlngPlanNo <> 0 Then
        Me.Caption = "�޸ķ���"
        txtPlanName.Text = mstrPlanName
        txtPlanDescription.Text = mstrDescription
    End If
End Sub
