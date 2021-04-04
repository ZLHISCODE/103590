VERSION 5.00
Begin VB.Form frmBloodReactionRecordSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2880
      TabIndex        =   3
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   345
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1100
   End
   Begin VB.TextBox txtTimer 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   200
      Left            =   705
      TabIndex        =   0
      Text            =   "10"
      Top             =   285
      Width           =   325
   End
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      Caption         =   "ÿ    �����Զ�ˢ�����������е�����"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3900
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      X1              =   0
      X2              =   4200
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   1000
      Y1              =   510
      Y2              =   510
   End
End
Attribute VB_Name = "frmBloodReactionRecordSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mstrPrivs As String
Private mblnOk As Boolean
Private mfrmMain As Object

Private mblnSetup As Boolean
Public Function ShowPara(ByVal frmMain As Object, Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mstrPrivs = strPrivs
    'mblnSetup = IsPrivs(mstrPrivs, "��������")
    mblnOk = False
    
    Set mfrmMain = frmMain
    '��ʼ��
    
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    If ExecuteCommand("��ȡ����") = False Then Exit Function
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    ShowPara = mblnOk
    
End Function

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rsSQL As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strValue As String
    
    On Error GoTo ErrHand
    
    Call SQLRecord(rsSQL)
    
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        
        '--------------------------------------------------------------------------------------------------------------
        Case "��ʼ����"
            
        '--------------------------------------------------------------------------------------------------------------
        Case "��ȡ����"
            Dim lngˢ�¼�� As Long
            
            lngˢ�¼�� = Val(gobjDatabase.GetPara("��Ϣ���Ѽ��", 2200, 1938, 0, Array(chk(0), txtTimer, Line1), mblnSetup))
            chk(0).Value = IIf(lngˢ�¼�� > 0, 1, 0)
            txtTimer.Enabled = chk(0).Value = 1
            txtTimer = IIf(chk(0).Value = 1, lngˢ�¼��, 10)
        '--------------------------------------------------------------------------------------------------------------
        Case "�������"
            If chk(0).Value = 1 Then
                Call gobjDatabase.SetPara("��Ϣ���Ѽ��", Val(txtTimer.Text), 2200, 1938, mblnSetup)
            Else
                Call gobjDatabase.SetPara("��Ϣ���Ѽ��", 0, 2200, 1938, mblnSetup)
            End If
            mblnOk = True
        End Select
    Next
    
    ExecuteCommand = True
    Exit Function
    '--------------------------------------------------------------------------------------------------------------
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub chk_Click(Index As Integer)
    
    If Visible = False Then Exit Sub
    cmdOK.Tag = "Changed"
    If Index = 0 Then
        txtTimer.Enabled = chk(0).Value = 1
        If Visible And txtTimer.Enabled Then txtTimer.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cmdOK.Tag <> "" Then
        If ExecuteCommand("�������") Then
            mblnOk = True
            cmdOK.Tag = ""
        End If
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbNo)
    End If
End Sub

Private Sub txtTimer_GotFocus()
    Call gobjControl.TxtSelAll(txtTimer)
End Sub

Private Sub txtTimer_KeyPress(KeyAscii As Integer)
    cmdOK.Tag = "Changed"
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    If Val(txtTimer.Text) > 99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub
