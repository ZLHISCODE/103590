VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmComments 
   BackColor       =   &H80000005&
   Caption         =   "�Զ��Ż�����"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11730
   FillColor       =   &H80000005&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   11730
   StartUpPosition =   1  '����������
   Begin RichTextLib.RichTextBox rtfComments 
      Height          =   7095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   12515
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmComments.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6960
      TabIndex        =   1
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "ִ��(&O)"
      Height          =   350
      Left            =   8160
      TabIndex        =   0
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label lblResult 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   90
   End
End
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrTask As String
Public Event UpdateStatus(ByVal strStatus As String)

Private Sub cmdCancel_Click()
    Dim strSQL As String, strQuery As String
    
    On Error GoTo errh
    strQuery = "ȷ����Ӧ�õ�ǰ�Զ��Ż�������" & vbNewLine
    If MsgBox(strQuery, vbYesNo + vbQuestion + vbDefaultButton1, "ȷ��") = vbNo Then Exit Sub
    
    strSQL = "Begin" & vbNewLine & _
                " Dbms_Sqltune.Drop_Tuning_Task('" & mstrTask & "');" & vbNewLine & _
                "end;"
    gcnOracle.Execute strSQL
    RaiseEvent UpdateStatus("������ȡ����")
    Unload Me
    Exit Sub
errh:
    ErrCenter
End Sub

Private Sub cmdExecute_Click()
    Dim strSQL As String, strQuery As String, rsData As ADODB.Recordset
    
    On Error GoTo errh
    strQuery = "ȷ��ҪӦ�õ�ǰ�Զ��Ż�������" & vbNewLine
    If MsgBox(strQuery, vbYesNo + vbQuestion + vbDefaultButton1, "ȷ��") = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    RaiseEvent UpdateStatus("�Զ��Ż���������ִ�С�")
    strSQL = "select Owner from dba_advisor_tasks  where TASK_NAME = [1]"
    Set rsData = OpenSQLRecord(strSQL, Me.Caption, mstrTask)
    
    strSQL = "Begin" & vbNewLine & _
                "dbms_sqltune.accept_sql_profile(task_name => '" & mstrTask & "', task_owner => '" & rsData!Owner & "', replace => TRUE,profile_type =>DBMS_SQLTUNE.PX_PROFILE);" & vbNewLine & _
                "end;"
    gcnOracle.Execute strSQL
    Screen.MousePointer = vbDefault
    RaiseEvent UpdateStatus("�Զ��Ż�����ִ����ɣ�")
    lblResult.Caption = "�Զ��Ż�����ִ����ɣ�"
    Unload Me
    Exit Sub
errh:
    lblResult.Caption = "ִ������ʧ�ܣ�����ݽ�����Ϣ�ֶ��Ż���"
    RaiseEvent UpdateStatus("ִ������ʧ�ܣ�����ݽ�����Ϣ�ֶ��Ż���")
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub Form_Load()
    Me.Icon = Nothing
End Sub

Private Sub Form_Resize()
    rtfComments.Move 0, 0, Me.ScaleWidth, Abs(Me.ScaleHeight - 735)
    cmdExecute.Top = rtfComments.Height + 225
    cmdExecute.Left = Me.ScaleWidth - cmdExecute.Width - 135
    cmdCancel.Top = cmdExecute.Top
    cmdCancel.Left = cmdExecute.Left - cmdCancel.Width - 60
    lblResult.Left = rtfComments.Left
    lblResult.Top = cmdCancel.Top
End Sub


Public Sub ShowFrm(strComments As String, strTask As String)
    rtfComments.Text = strComments
    mstrTask = strTask
    Me.Show 1

End Sub

