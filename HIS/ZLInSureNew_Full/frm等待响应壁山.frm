VERSION 5.00
Begin VB.Form frm�ȴ���Ӧ��ɽ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ȴ���Ӧ..."
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   ControlBox      =   0   'False
   Icon            =   "frm�ȴ���Ӧ��ɽ.frx":0000
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
      Picture         =   "frm�ȴ���Ӧ��ɽ.frx":000C
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
Attribute VB_Name = "frm�ȴ���Ӧ��ɽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrBillNo As String     '���㵥��
Private mintClass As Integer     '����
Private mintChargeNo As Integer  '���õ����
Private mblnReturn As Boolean   '���ؽ��

Private Sub cmdCancel_Click()
    TimeSearch.Enabled = True
    mblnReturn = False
    Unload Me
End Sub

Public Function Result(int��� As Integer, strBill_no As String, Optional intNo As Integer) As Boolean
'���ܣ��ж�Ѱ�ҵĽ��
'������int���  1���Ǽ�  2������
    mintClass = int���
    mstrBillNo = strBill_no
    mintChargeNo = intNo
    Me.Show 1
    Result = mblnReturn
End Function

Private Sub Form_Activate()
    Dim strSql As String, rs��ɽ As New ADODB.Recordset
    If mstrBillNo = "" Then Exit Sub
    '��ѯ�Ƿ��з��صĽ��
    If mintClass = 1 Then
        strSql = "Select Request_Result from Check_bill_request where " & _
                "Bill_no = '" & mstrBillNo & "' and App_code = '" & _
                Mid(gstrҽԺ����, 1, 4) & "'"
    Else
        strSql = "select Request_Result from check_item_request where " & _
                "Bill_no = '" & mstrBillNo & "' and App_code = '" & _
                Mid(gstrҽԺ����, 1, 4) & "' and Charge_item_no = '" & CStr(mintChargeNo) & "'"
    End If
    '
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
    If Not IsNull(rs��ɽ("Request_Result")) And rs��ɽ("Request_Result") <> 0 Then
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

Private Sub TimeAvi_Timer()
    Static i As Long
    TimeSearch.Enabled = True
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub

Private Sub TimeSearch_Timer()
    Dim strSql As String, rs��ɽ As New ADODB.Recordset, lngErrLine As Long
    
    If mstrBillNo = "" Then Exit Sub
    '��ѯ�Ƿ��з��صĽ��
    On Error GoTo errHandle
    If mintClass = 1 Then
        strSql = "Select Request_Result from Check_bill_request where " & _
                "Bill_no = '" & mstrBillNo & "' and App_code = '" & _
                Mid(gstrҽԺ����, 1, 4) & "'": lngErrLine = 1
    Else
        strSql = "select Request_Result from check_item_request where " & _
                "Bill_no = '" & mstrBillNo & "' and App_code = '" & _
                Mid(gstrҽԺ����, 1, 4) & "' and Charge_item_no = '" & CStr(mintChargeNo) & "'": lngErrLine = 2
    End If
    '
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close: lngErrLine = 3
    rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly: lngErrLine = 4
    If Not IsNull(rs��ɽ("Request_Result")) Then
        mblnReturn = True: lngErrLine = 5
        TimeSearch.Enabled = False: lngErrLine = 6
        Unload Me
    Else
        mblnReturn = False: lngErrLine = 7
    End If
    Exit Sub
errHandle:
    MsgBox "��[�ȴ���Ӧ]�����У�[TimeSearch_Timer]�¼���" & lngErrLine & "�з�������", vbExclamation, "����"
    
End Sub
