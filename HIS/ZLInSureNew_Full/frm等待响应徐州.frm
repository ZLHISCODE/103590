VERSION 5.00
Begin VB.Form frm�ȴ���Ӧ���� 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ȴ���Ӧ..."
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frm�ȴ���Ӧ����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      Picture         =   "frm�ȴ���Ӧ����.frx":000C
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1170
      Width           =   5325
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.Timer TimeSearch 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   720
   End
   Begin VB.Timer TimeAvi 
      Interval        =   50
      Left            =   2040
      Top             =   690
   End
   Begin VB.Label LblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  ���ύ�������ڵȴ�ҽ����������Ӧ..."
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1350
      TabIndex        =   2
      Top             =   510
      Width           =   3510
   End
End
Attribute VB_Name = "frm�ȴ���Ӧ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrBillNo As String     '���㵥��
Private mblnReturn As Boolean   '���ؽ��
Private mstrReturn As String

Private Sub cmdCancel_Click()
    TimeSearch.Enabled = False
    mblnReturn = False
    Unload Me
End Sub

Public Function Result(strTransNO As String, strReturn As String) As Boolean
'���ܣ��ж�Ѱ�ҵĽ��
'������int���  1���Ǽ�  2������
    mstrBillNo = strTransNO
    Me.Show 1
    strReturn = mstrReturn
    Result = mblnReturn
End Function

Private Sub Form_Activate()
    Dim strSql As String, rs���� As New ADODB.Recordset
    If mstrBillNo = "" Then Exit Sub
    '��ѯ�Ƿ��з��صĽ��
    strSql = "Select * from ins_result Where transerial='" & mstrBillNo & "'"
    If rs����.State = adStateOpen Then rs����.Close
    rs����.Open strSql, gcn����, adOpenStatic, adLockReadOnly
    If rs����.EOF Then
        mblnReturn = False
        TimeSearch.Enabled = True
    ElseIf Nvl(rs����!tranflag, 9) = 0 Then
        mstrReturn = Nvl(rs����!Result, "")
        mblnReturn = True
        TimeSearch.Enabled = False
        Unload Me
    ElseIf Nvl(rs����!tranflag, 9) = 8 Then
        MsgBox rs����!info, vbInformation, "ҽ������"
        mblnReturn = True
        TimeSearch.Enabled = False
        Unload Me
    Else
        mblnReturn = False
        TimeSearch.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    mblnReturn = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnReturn = False Then Cancel = 1
End Sub

Private Sub TimeAvi_Timer()
    Static i As Long
    TimeSearch.Enabled = True
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub

Private Sub TimeSearch_Timer()
    Dim strSql As String, rs���� As New ADODB.Recordset
    
    If mstrBillNo = "" Then Exit Sub
    '��ѯ�Ƿ��з��صĽ��
'    strSql = "Select * from ins_tranask Where transerial='" & mstrBillNo & "'"
'    If rs����.State = adStateOpen Then rs����.Close
'    rs����.Open strSql, gcn����, adOpenStatic, adLockReadOnly
'    If Nvl(rs����!tranflag, 9) = 8 Then
'        MsgBox "���״���ʧ��", vbInformation, gstrSysName
'        mblnReturn = True
'        TimeSearch.Enabled = False
'        Unload Me
'    End If
    
    strSql = "Select * from ins_result Where transerial='" & mstrBillNo & "'"
    If rs����.State = adStateOpen Then rs����.Close
    rs����.Open strSql, gcn����, adOpenStatic, adLockReadOnly
    
    If rs����.EOF Then
        mblnReturn = False
        TimeSearch.Enabled = True
    ElseIf Nvl(rs����!tranflag, 9) = 0 Then
        mstrReturn = Nvl(rs����!Result, "")
        If mstrReturn = "01" Then
            MsgBox rs����!info, vbInformation, "ҽ������"
        End If
        WriteInfo "���״���:" & rs����!info
        mblnReturn = True
        TimeSearch.Enabled = False
        Unload Me
    ElseIf Nvl(rs����!tranflag, 9) = 8 Then
        MsgBox rs����!info, vbInformation, "ҽ������"
        mblnReturn = True
        TimeSearch.Enabled = False
        Unload Me
    Else
        mblnReturn = False
        TimeSearch.Enabled = True
    End If
End Sub
