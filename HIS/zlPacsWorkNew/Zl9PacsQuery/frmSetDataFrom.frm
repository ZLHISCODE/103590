VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSetDataFrom 
   Caption         =   "�༭����"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7440
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
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin RichTextLib.RichTextBox rctData 
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8916
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"frmSetDataFrom.frx":6852
   End
   Begin VB.CommandButton cmdExample 
      Caption         =   "��������"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   5475
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsertPar 
      Caption         =   "�������(&I)"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5475
      Width           =   1095
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "��ѯ��֤"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   5475
      Width           =   1095
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   375
      Left            =   5040
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
      TabIndex        =   2
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


Public Function ShowSqlFromWindow(ByVal strSqlFrom As String, strPara As String, lngCaption As Long, IsEnabled As Boolean, ByVal bytSize As Byte, owner As Object) As String
    ShowSqlFromWindow = strSqlFrom
    
    Me.mblnIsOK = False
    Me.rctData.Text = strSqlFrom
    Me.rctData.Locked = Not IsEnabled
    Me.cmdSure.Enabled = IsEnabled
    Me.cmdInsertPar.Enabled = IsEnabled
    Me.cmdInsertPar.Caption = "�������"
    
    Call SetFontSize(bytSize)
    Select Case lngCaption
        Case 1
            Me.Caption = "Ĭ��ֵ����"
            lblHint.Caption = "Ĭ��ֵ�����Ƕ�¼����ĿĬ��ȡֵ��"
            cmdVerify.Visible = False
        Case 2
            Me.Caption = "������Դ����"
            lblHint.Caption = "������Դ�����У����ʹ��sql��䣬�����ʹ��ǰ��" & vbCrLf & "��Ŀ��¼��ֵ��Ϊ��ѯ������������"
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
        If Len(Trim(rctData.Text)) = 0 Or InStr(UCase(Trim(rctData.Text)), UCase("select")) = 0 Then
            cmdVerify.Enabled = False
        Else
            cmdVerify.Enabled = True
        End If
    Else
        cmdVerify.Enabled = False
    End If

    If Me.rctData.Locked Then
        Me.rctData.BackColor = &H8000000F
    Else
        Me.rctData.BackColor = &H80000005
    End If
    
    mstrPara = strPara
    Call Me.Show(1, owner)
    
    If Me.mblnIsOK Then
        ShowSqlFromWindow = Me.rctData.Text
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
        strPar = frmPar.ShowParameterWindow(False, Me, mstrPara, 1)
        If strPar <> "" Then
            rctData.SelText = strPar
        End If
        
        Set frmPar = Nothing
    ElseIf cmdInsertPar.Caption = "��������" Then
        rctData.Text = "Function CustomFilterScript(rsStudyData, strFilterWhere)" & vbCrLf & _
                            "   " & "set CustomFilterScript = ����ֵ" & vbCrLf & _
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
    
    strErr = SqlVerify(rctData.Text)
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
    
    rctData.Top = lblHint.Top + lblHint.Height + 60
    rctData.Width = Me.ScaleWidth - rctData.Left * 2
    rctData.Height = Me.ScaleHeight - rctData.Top - cmdSure.Height - 120
    
    cmdInsertPar.Top = rctData.Top + rctData.Height + 60
    
    cmdVerify.Left = cmdInsertPar.Left + cmdInsertPar.Width + 60
    cmdVerify.Top = cmdInsertPar.Top
    
    cmdExample.Top = cmdInsertPar.Top
    
    cmdCancel.Left = rctData.Width + rctData.Left - cmdCancel.Width
    cmdCancel.Top = cmdInsertPar.Top
    
    cmdSure.Left = cmdCancel.Left - 60 - cmdSure.Width
    cmdSure.Top = cmdInsertPar.Top
End Sub

Private Sub rctData_Change()
    On Error GoTo errHandle
    
    If Len(Trim(rctData.Text)) = 0 Or InStr(UCase(Trim(rctData.Text)), UCase("select")) = 0 Then
        cmdVerify.Enabled = False
    Else
        cmdVerify.Enabled = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim lngCmdHeight As Long
    Dim lngCmdWithd As Long
    
    If bytFontSize = 9 Then
        lngCmdHeight = 350
        lngCmdWithd = 1100
    ElseIf bytFontSize = 12 Then
        lngCmdHeight = 385
        lngCmdWithd = 1300
    ElseIf bytFontSize = 15 Then
        lngCmdHeight = 420
        lngCmdWithd = 1500
    End If
    
    lblHint.FontSize = bytFontSize
    rctData.Font.Size = bytFontSize
    
    cmdCancel.FontSize = bytFontSize
    cmdCancel.Height = lngCmdHeight
    cmdCancel.Width = lngCmdWithd
    
    cmdInsertPar.FontSize = bytFontSize
    cmdInsertPar.Height = lngCmdHeight
    cmdInsertPar.Width = lngCmdWithd
    
    cmdExample.FontSize = bytFontSize
    cmdExample.Height = lngCmdHeight
    cmdExample.Width = lngCmdWithd
    
    cmdSure.FontSize = bytFontSize
    cmdSure.Height = lngCmdHeight
    cmdSure.Width = lngCmdWithd
    
    cmdVerify.FontSize = bytFontSize
    cmdVerify.Height = lngCmdHeight
    cmdVerify.Width = lngCmdWithd
End Sub
