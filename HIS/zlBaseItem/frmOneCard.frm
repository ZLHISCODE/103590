VERSION 5.00
Begin VB.Form frmOneCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "һ��ͨ����"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5805
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1890
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   4320
      TabIndex        =   13
      Top             =   -120
      Width           =   30
   End
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "����"
      Text            =   "11"
      Top             =   240
      Width           =   1125
   End
   Begin VB.TextBox txtOrgCode 
      Height          =   300
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   7
      Tag             =   "����"
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4530
      TabIndex        =   10
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4530
      TabIndex        =   11
      Top             =   690
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "����"
      Top             =   600
      Width           =   2925
   End
   Begin VB.ComboBox cboPayType 
      Height          =   300
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1050
      Width           =   2025
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4530
      TabIndex        =   12
      Top             =   1560
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&E)"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1950
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ҽԺ����(&O)"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1530
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   660
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "���(&U)"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   300
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "���㷽ʽ(&P)"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   1110
      Width           =   990
   End
End
Attribute VB_Name = "frmOneCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbytInFun As Byte '0-����,1-�޸�

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, "frmOneCard", Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    
    If cboPayType.ListIndex = -1 Then
        MsgBox "��ѡ����㷽ʽ!", vbInformation, gstrSysName
        cboPayType.SetFocus
        Exit Sub
    End If
    If txtName.Text = "" Then
        MsgBox "������һ��ͨ�ӿ�����!", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    If txtOrgCode.Text = "" Then
        MsgBox "������ҽԺ����!", vbInformation, gstrSysName
        txtOrgCode.SetFocus
        Exit Sub
    End If
    
    If zlCommFun.ActualLen(txtName.Text) > txtName.MaxLength Then
        MsgBox "���Ʋ��ܳ���" & txtName.MaxLength & "���ַ�!", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    If zlCommFun.ActualLen(txtOrgCode.Text) > txtOrgCode.MaxLength Then
        MsgBox "ҽԺ���벻�ܳ���" & txtOrgCode.MaxLength & "���ַ�!", vbInformation, gstrSysName
        txtOrgCode.SetFocus
        Exit Sub
    End If

    '�ò���һ��Ϊϵͳ����Ա����,���Ժ��Բ�������, �����ݽṹ����
    strSQL = "Zl_һ��ͨĿ¼_Update(" & txtNO.Text & ",'" & txtName.Text & "','" & cboPayType.Text & "','" & _
            txtOrgCode.Text & "'," & cbo����.ListIndex & "," & mbytInFun & ")"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, App.ProductName)
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    txtName.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not Me.ActiveControl Is cmdOK Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Function GetPayType() As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select ����,���� From ���㷽ʽ Where ����=7"
    On Error GoTo errH
    Set GetPayType = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOneCardMaxNO() As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(max(���),0) ��� From һ��ͨĿ¼"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    GetOneCardMaxNO = Val(rsTmp!���) + 1
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowMe(objParent As Form, Optional intNO As Integer, Optional strName As String, _
    Optional strPayType As String, Optional strOrgCode As String, Optional intState As Integer)
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetPayType
    If rsTmp.RecordCount = 0 Then
        MsgBox "û���ҵ�����Ϊһ��ͨ�Ľ��㷽ʽ,���ȵ�[���㷽ʽ����]�����á�", vbInformation
        Exit Sub
    End If
    Call zlControl.CboAddData(cboPayType, rsTmp, True)
    
    With Me.cbo����
        .Clear
        .AddItem "ͣ��", 0
        .AddItem "����:���漰�ۿ�", 1
        .AddItem "����:��׼һ��ͨ", 2
        .ListIndex = 0
    End With
    
    If mbytInFun = 0 Then
        cboPayType.ListIndex = 0
        txtNO.Text = GetOneCardMaxNO
    Else
        txtNO.Text = intNO
        txtName.Text = strName
        Call zlControl.CboLocate(cboPayType, strPayType)
        txtOrgCode.Text = strOrgCode
        cbo����.ListIndex = intState
    End If
    
    Me.Show 1, objParent
End Sub

Public Sub DelOneCardRec(intNO As Integer)
    Dim strSQL As String
    
    strSQL = "Zl_һ��ͨĿ¼_Update(" & intNO & ",null,null,null,null,2)"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, App.ProductName)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    txtName.Text = Trim(txtName.Text)
End Sub

Private Sub txtOrgCode_Change()
    txtOrgCode.Text = Trim(txtOrgCode.Text)
End Sub
