VERSION 5.00
Begin VB.Form frmExitPsw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   1965
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5580
   Icon            =   "frmExitPsw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer tmr 
      Interval        =   30000
      Left            =   4530
      Top             =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   1830
      Left            =   45
      TabIndex        =   4
      Top             =   60
      Width           =   4215
      Begin VB.TextBox TXT���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1335
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   585
         Width           =   1920
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Left            =   3240
         TabIndex        =   6
         Top             =   450
         Width           =   900
         _extentx        =   1588
         _extenty        =   1005
         caption         =   "�����"
         backcolor       =   16777215
         fontsize        =   10.5
         autosize        =   0   'False
         buttonheight    =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000FF&
         Height          =   630
         Left            =   1440
         TabIndex        =   5
         Top             =   1095
         Width           =   2640
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   885
         TabIndex        =   0
         Top             =   645
         Width           =   360
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   180
         Picture         =   "frmExitPsw.frx":000C
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4395
      TabIndex        =   3
      Top             =   615
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4395
      TabIndex        =   2
      Top             =   135
      Width           =   1100
   End
End
Attribute VB_Name = "frmExitPsw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean

Private Function OraData(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim StrSQL As String
    Dim strError As String
    Dim cnOracle As New ADODB.Connection
    
    On Error Resume Next
    Err = 0
    DoEvents
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                lbl.Caption = "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��"
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                lbl.Caption = "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������"
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                lbl.Caption = "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������"
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                lbl.Caption = "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�"
            Else
                lbl.Caption = "�����û�������������ָ�������޷�ע�ᡣ"
                
            End If
            
            OraData = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand

    OraData = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraData = False
    Err = 0
End Function

Public Function ShowPsw(ByVal frmMain As Object, Optional blnHasSoftKeyBoard As Boolean = False) As Boolean
    mblnOK = False
    Me.UsrCmd.ShowPicture = False
    Me.UsrCmd.Visible = blnHasSoftKeyBoard
    Me.Show 1, frmMain
    ShowPsw = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strPassword As String
    
    strPassword = TXT����.Text
    If Not OraData(gstrServerName, gstrDbUser, IIf(UCase(gstrDbUser) = "SYS" Or UCase(gstrDbUser) = "SYSTEM", strPassword, TranPasswd(strPassword))) Then
'        TXT����.Text = ""
        If TXT����.Enabled Then TXT����.SetFocus
        Exit Sub
    End If
    
    mblnOK = True
    
    Unload Me

End Sub

Private Sub tmr_Timer()
    If Me.UsrCmd.Visible Then Unload frmCheckQueryPass
    Unload Me
End Sub

Private Sub TXT����_Change()
    lbl.Caption = ""
End Sub

Private Sub TXT����_GotFocus()
    zlControl.TxtSelAll TXT����
End Sub


Private Sub UsrCmd_CommandClick()
    If frmCheckQueryPass.GetPwd(Me) = False Then Unload Me: Exit Sub
    TXT����.Text = frmCheckQueryPass.mstrPass
    Call cmdOK_Click
End Sub

