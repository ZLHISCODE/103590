VERSION 5.00
Begin VB.Form frmQCDiscardEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����д����ԭ��"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5205
   Icon            =   "frmQCDiscardEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2970
      TabIndex        =   2
      Top             =   2775
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1110
      TabIndex        =   1
      Top             =   2775
      Width           =   1100
   End
   Begin VB.TextBox txtԭ�� 
      Height          =   2580
      Left            =   135
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   105
      Width           =   4920
   End
End
Attribute VB_Name = "frmQCDiscardEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrԭ�� As String
Private mblnOK As Boolean

Public Function ShowMe(ByVal lngID As Long, ByRef strԭ�� As String, ByVal frmMain As Form) As Boolean
    Dim rsTmp As adodb.Recordset, strSQL As String
    mblnOK = False
    mstrԭ�� = ""
    If lngID > 0 Then
        strSQL = "select ԭ�� From �������ñ��� Where ���ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        Do Until rsTmp.EOF
            mstrԭ�� = Trim("" & rsTmp!ԭ��)
            rsTmp.MoveNext
        Loop
        Me.txtԭ��.Text = mstrԭ��
        Me.Show vbModal, frmMain
        
        If mblnOK Then
            strԭ�� = mstrԭ��
            ShowMe = mblnOK
        End If
    End If
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    mstrԭ�� = Replace(Me.txtԭ��.Text, "'", "")
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
'    Me.txtԭ��.Text = mstrԭ��
End Sub
