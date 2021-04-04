VERSION 5.00
Begin VB.Form frmSaveAs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ʵ�ģ��"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmSaveAs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraSave 
      Caption         =   "�¼��ʵ�ģ����Ϣ"
      Height          =   1485
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4395
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   960
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "����"
         Top             =   840
         Width           =   3165
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   960
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "����"
         Top             =   420
         Width           =   1725
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   480
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3390
      TabIndex        =   6
      Top             =   1890
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2190
      TabIndex        =   5
      Top             =   1890
      Width           =   1100
   End
End
Attribute VB_Name = "frmSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrID As String
Dim mstr���� As String
Dim mstr���� As String
Dim mblnOK As Boolean
Dim mblnSave  As Boolean       '����������ƺ��Ƿ���Ҫ��������
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save���ʵ�() = False Then Exit Sub
    
    Unload Me
End Sub

Private Function IsValid() As Boolean
'����:���������йؼ��ʵ��������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 1
        strTemp = Trim(txtEdit(i).Text)
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox "���������ݲ��ܳ���" & Int(txtEdit(i).MaxLength / 2) & "������" & "��" & txtEdit(i).MaxLength & "����ĸ��", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
    Next
    If Len(Trim(txtEdit(0).Text)) = 0 Then
        txtEdit(0).Text = ""
        MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(0).SetFocus
        Exit Function
    End If
    If Len(Trim(txtEdit(1).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(1).Text = ""
        txtEdit(1).SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function Save���ʵ�() As Boolean
'����:����༭�����ݵ����ʵ�����
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim lngID As Long
    On Error GoTo errHandle
    
    lngID = zlDatabase.GetNextId("�շѼ��ʵ�")
    mstr���� = txtEdit(0).Text
    mstr���� = txtEdit(1).Text
    
    If mblnSave = True Then
        gstrSQL = "zl_�շѼ��ʵ�_SaveAs('" & lngID & _
            "','" & mstr���� & "','" & mstr���� & "','" & mstrID & "')"
            
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
        
    mblnOK = True
    mblnChange = False
    mstrID = CStr(lngID)
    Save���ʵ� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ���ģ��(strID As String, str���� As String, str���� As String, Optional ByVal blnSave As Boolean = True) As Boolean
'����:��������õļ��ʵ������ڽ���ͨѶ�ĳ���
'����:ʵ��������Ϊ����ֵ
'     strID �ڴ���ʱ�ǲ���ģ���ID������ʱ����Ϊ����ģ���ID
    mblnChange = False
    mblnSave = blnSave
    mblnOK = False
    mstrID = strID
    frmSaveAs.Show vbModal
    
    If mblnOK = True Then
        strID = mstrID
        str���� = mstr����
        str���� = mstr����
    End If
    ���ģ�� = mblnOK
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
    End If
End Sub


