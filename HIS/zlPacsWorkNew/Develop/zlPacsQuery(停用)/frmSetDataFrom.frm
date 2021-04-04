VERSION 5.00
Begin VB.Form frmSetDataFrom 
   Caption         =   "�༭����"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetDataFrom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7455
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdExample 
      Caption         =   "��������"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   5475
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsertPar 
      Caption         =   "�������(&I)"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5475
      Width           =   1095
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "��ѯ��֤"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   5475
      Width           =   1095
   End
   Begin VB.TextBox txtVulue 
      Height          =   5040
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   7215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   375
      Left            =   5085
      TabIndex        =   1
      Top             =   5475
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   5475
      Width           =   1095
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmSetDataFrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnIsOK As Boolean
Private mstrPara As String


Public Function ShowSqlFromWindow(ByVal strSqlFrom As String, strPara As String, lngCaption As Long, IsEnabled As Boolean, owner As Object) As String
    ShowSqlFromWindow = strSqlFrom
    
    Me.mblnIsOK = False
    Me.txtVulue.Text = strSqlFrom
    Me.txtVulue.Locked = Not IsEnabled
    Me.cmdSure.Enabled = IsEnabled
    Me.cmdInsertPar.Enabled = IsEnabled
    Me.cmdInsertPar.Caption = "�������"
    
    Select Case lngCaption
        Case 1
            Me.Caption = "Ĭ��ֵ����"
            lblHint.Caption = "Ĭ��ֵ�����Ƕ�¼����ĿĬ��ȡֵ��"
            cmdVerify.Visible = False
        Case 2
            Me.Caption = "������Դ����"
            lblHint.Caption = "������Դ�����У����ʹ��sql��䣬�����ʹ��ǰ����Ŀ��¼��ֵ��Ϊ��ѯ������������"
        Case 3
            Me.Caption = "����ת������"
            lblHint.Caption = "����ת�����õĸ�ʽ���磺1-��;2-Ů"
            cmdInsertPar.Visible = False
            cmdVerify.Visible = False
        Case 4
            Me.Caption = "�Զ�����˽ű�"
            lblHint.Caption = "�Զ�����˽ű���Ҫ���ڸ��ӵ����ݹ��ˡ�"
            cmdInsertPar.Caption = "��������"
            cmdVerify.Visible = False
    End Select
    
    If IsEnabled Then
        If Len(Trim(txtVulue.Text)) = 0 Or InStr(UCase(Trim(txtVulue.Text)), UCase("select")) = 0 Then
            cmdVerify.Enabled = False
        Else
            cmdVerify.Enabled = True
        End If
    Else
        cmdVerify.Enabled = False
    End If

    If Me.txtVulue.Locked Then
        Me.txtVulue.BackColor = &H8000000F
    Else
        Me.txtVulue.BackColor = &H80000005
    End If
    
    mstrPara = strPara
    Call Me.Show(1, owner)
    
    If Me.mblnIsOK Then
        ShowSqlFromWindow = Me.txtVulue.Text
    Else
        ShowSqlFromWindow = strSqlFrom
    End If

End Function

Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    mblnIsOK = False
    
    Call Me.Hide
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub


Private Sub cmdInsertPar_Click()
'�������
    Dim strPar As String
    Dim frmPar As New frmSetPara
    
    On Error GoTo errHandle
    
    If cmdInsertPar.Caption = "�������" Then
        strPar = frmPar.ShowParameterWindow(False, Me, mstrPara)
        If strPar <> "" Then
            txtVulue.SelText = strPar
        End If
        
        Set frmPar = Nothing
    ElseIf cmdInsertPar.Caption = "��������" Then
        txtVulue.Text = "Function CustomFilterScript(rsStudyData, strFilterWhere)" & vbCrLf & _
                            "   " & "CustomFilterScript = ����ֵ" & vbCrLf & _
                        "End Function"
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cmdSure_Click()
    On Error GoTo errHandle
    
    mblnIsOK = True
    Call Me.Hide
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cmdVerify_Click()
    Dim strErr As String
    
    On Error GoTo errHandle
    
    strErr = SqlVerify(txtVulue.Text)
    If Len(strErr) = 0 Then
        MsgBox "��֤�ɹ���", vbInformation, Me.Caption
    Else
        MsgBox "��֤ʧ�ܣ�ԭ��Ϊ��" & strErr, vbInformation, Me.Caption
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    mblnIsOK = False
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    txtVulue.Width = Me.ScaleWidth - txtVulue.Left * 2
    txtVulue.Height = Me.ScaleHeight - txtVulue.Top - cmdSure.Height - 120
    
    cmdInsertPar.Top = txtVulue.Top + txtVulue.Height + 60
    cmdVerify.Top = cmdInsertPar.Top
    
    cmdCancel.Left = txtVulue.Width + txtVulue.Left - cmdCancel.Width
    cmdCancel.Top = cmdInsertPar.Top
    
    cmdSure.Left = cmdCancel.Left - 60 - cmdSure.Width
    cmdSure.Top = cmdInsertPar.Top
End Sub

Private Sub txtVulue_Change()
    On Error GoTo errHandle
    
    If Len(Trim(txtVulue.Text)) = 0 Or InStr(UCase(Trim(txtVulue.Text)), UCase("select")) = 0 Then
        cmdVerify.Enabled = False
    Else
        cmdVerify.Enabled = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub
