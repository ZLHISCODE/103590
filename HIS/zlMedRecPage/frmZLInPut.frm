VERSION 5.00
Begin VB.Form frmZLInPut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������������̬ѧ����"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4005
   Icon            =   "frmZLInPut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4005
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   1100
   End
   Begin VB.PictureBox picSentence 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   2835
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   2865
      Begin VB.TextBox txtSentence 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   15
         TabIndex        =   0
         Top             =   30
         Width           =   2610
      End
      Begin VB.Image imgSentence 
         Height          =   210
         Left            =   2640
         Picture         =   "frmZLInPut.frx":6852
         ToolTipText     =   "�밴 * �ż�ѡ��"
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1100
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmZLInPut.frx":6D7C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Caption         =   "   ���[���ǳ�ԭ�Զ�������]Ϊ������ϣ�������������̬ѧ���룡"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   315
      Width           =   3015
   End
End
Attribute VB_Name = "frmZLInPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrCaption As String
Private mrsOutPut As Recordset

Public Function ShowMe(objForm As Object, ByVal strCaption As String, ByRef rsOutPut As Recordset) As Boolean
    mstrCaption = strCaption
    mblnOK = False
    Set mrsOutPut = Nothing
    Me.Show 1, objForm
    Set rsOutPut = mrsOutPut
    ShowMe = mblnOK
End Function

Private Sub cmdCancle_Click()
    Set mrsOutPut = Nothing
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If mrsOutPut Is Nothing Then
        MsgBox "��¼����ȷ��������̬ѧ����!", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    lblTitle.Caption = mstrCaption
End Sub

Private Sub imgSentence_Click()
    Set mrsOutPut = zlDatabase.ShowILLSelect(Me, "M", gclsPros.��Ժ����ID, gclsPros.Sex, False, , , gclsPros.SysNo)
    If Not mrsOutPut Is Nothing Then
        txtSentence.Text = mrsOutPut!���� & ""
        txtSentence.Tag = txtSentence.Text
    End If
End Sub

Private Sub txtSentence_GotFocus()
    zlControl.TxtSelAll txtSentence
End Sub

Private Sub txtSentence_KeyPress(KeyAscii As Integer)
    Dim strInput As String
    Dim blnCancel As Boolean
    Dim strSql As String
    Dim vPoint As POINTAPI
    
    If KeyAscii = vbKeyReturn Then
            
        strInput = UCase(txtSentence.Text)
        'B-��ҽ�������룬7-�����ж���Y-�����ж����ⲿԭ��6-������ϣ�M-������̬ѧ���룻������ϣ�D-ICD-10��������
        strSql = GetMedInputSQL(0, strInput, gclsPros.Sex)
        vPoint = GetCoordPos(txtSentence.Container.hwnd, txtSentence.Left, txtSentence.Top)
        Set mrsOutPut = zlDatabase.ShowSQLSelect(Me, strSql, 0, "��������", _
            False, "", "", False, False, True, vPoint.X, vPoint.Y + 15, txtSentence.Height, blnCancel, False, True, _
            strInput & "%", gclsPros.LikeString & strInput & "%", "M", gclsPros.Sex, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
        If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
            blnCancel = True
            txtSentence.Text = txtSentence.Tag
            zlControl.TxtSelAll txtSentence
        Else
            '���������뷽ʽ
            If mrsOutPut Is Nothing Then
                MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                blnCancel = True
                zlControl.TxtSelAll txtSentence
            Else
                txtSentence.Text = mrsOutPut!���� & ""
                txtSentence.Tag = txtSentence.Text
                cmdOk.SetFocus
            End If
        End If
    End If
End Sub

