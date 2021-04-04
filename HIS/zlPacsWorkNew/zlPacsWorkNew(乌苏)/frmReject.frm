VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���沵��"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8655
   Icon            =   "frmReject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin RichTextLib.RichTextBox rtbRejectHistory 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7858
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReject.frx":0AE2
   End
   Begin VB.CommandButton CmdCancle 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7440
      TabIndex        =   9
      Top             =   4920
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6120
      TabIndex        =   8
      ToolTipText     =   "����(F2)"
      Top             =   4920
      Width           =   1125
   End
   Begin VB.TextBox txtRejectUser 
      BackColor       =   &H8000000F&
      Height          =   300
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4440
      Width           =   3735
   End
   Begin MSComCtl2.DTPicker dtpRejectDate 
      Height          =   300
      Left            =   4800
      TabIndex        =   5
      Top             =   3720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd hh:mm"
      Format          =   129957891
      CurrentDate     =   41074
   End
   Begin RichTextLib.RichTextBox rtbReason 
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1826
      _Version        =   393217
      TextRTF         =   $"frmReject.frx":0B7F
   End
   Begin VB.Label Label2 
      Caption         =   "�����ˣ�"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "����ʱ�䣺"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "����ԭ��"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "��ʷ��¼��"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmReject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsOk As Boolean

Private mlngAdviceID As Long
Private mlngReportID As Long

Public Sub ShowRejectWindow(ByVal lngAdviceID As Long, ByVal lngReportID As Long, _
    owner As Object)
'��ʾ���沵�ش���
    mlngAdviceID = lngAdviceID
    mlngReportID = lngReportID
    IsOk = False
    
    rtbRejectHistory.Height = Me.ScaleHeight - cmdOK.Height - Label1.Height * 2 - rtbReason.Height - 480
    
    dtpRejectDate.value = zlDatabase.Currentdate
    txtRejectUser.Text = UserInfo.����
    
    Me.Caption = "���沵��"
    
    Call LoadRejectHistory(lngAdviceID, lngReportID)
    
    Me.Show 1, owner
End Sub

Public Sub ShowRejectHistory(ByVal lngAdviceID As Long, ByVal lngReportID As Long, _
    owner As Object)
'��ʾ������ʷ��¼
    mlngAdviceID = lngAdviceID
    mlngReportID = lngReportID
    
    IsOk = False
    
    rtbRejectHistory.Height = Me.ScaleHeight - cmdOK.Height - Label1.Height - 360
    
    cmdOK.Visible = False
    CmdCancle.Caption = "ȷ��(&O)"
    
    Me.Caption = "������ʷ"
    
    Call LoadRejectHistory(lngAdviceID, lngReportID)
    
    Me.Show 1, owner
End Sub

Private Sub LoadRejectHistory(ByVal lngAdviceID As Long, ByVal lngReportID As Long)
'���벵����ʷ��¼
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strHistory As String
    Dim strFormats As String
    
    strSQL = "select ��������,����ʱ��,������ from Ӱ�񱨸沵�� where ҽ��ID=[1] and ����Id=[2] order by ����ʱ�� Desc"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, lngReportID)
    
    rtbRejectHistory.Text = ""
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    strFormats = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\f0\fs20 "
    
    While Not rsData.EOF
        If strHistory <> "" Then strHistory = strHistory & vbCrLf
        
        strHistory = strHistory & "\b�������ɣ�" & "\par" & "\b0    " & Nvl(rsData!��������) & "\par" & _
                                    "\b                                    �����ˣ�\b0" & Nvl(rsData!������) & "\par" & _
                                    "\b                                    ����ʱ�䣺\b0 " & Format(Nvl(rsData!����ʱ��), "yyyy-mm-dd hh:mm:ss") & _
                                    "\par" & "----------------------------------------------------------------------" & "\par"
        
'        strHistory = strHistory & "�������ɣ�" & vbCrLf & "    " & Nvl(rsData!��������) & vbCrLf & _
'                                    "                                          �����ˣ�" & Nvl(rsData!������) & vbCrLf & _
'                                    "                                    ����ʱ�䣺" & Nvl(rsData!����ʱ��) & _
'                                    vbCrLf & "-------------------------------------------------------------------" & vbCrLf

'        strHistory = strHistory & Nvl(rsData!����ʱ��) & "    �����ˣ�" & Nvl(rsData!������) & "    " & vbCrLf & _
'                    "    " & Nvl(rsData!��������) & vbCrLf & vbCrLf
        
        rsData.MoveNext
    Wend
    
    rtbRejectHistory.TextRTF = strFormats & strHistory & "\par}"
End Sub

Private Sub CmdCancle_Click()
On Error GoTo ErrHandle
    IsOk = False
    Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHandle
    Dim strSQL As String
    
    If rtbReason.Text = "" Then
        MsgBoxD Me, "��¼�벵�����ɣ��Ա�˱���ҽ���˽ⱻ���ص�ԭ��", vbInformation, Me.Caption
        Exit Sub
    End If
    
    strSQL = "ZL_Ӱ�񱨸沵��(" & mlngAdviceID & "," & mlngReportID & ",'" & rtbReason.Text & "'," & To_Date(dtpRejectDate.value) & ",'" & txtRejectUser.Text & "')"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    IsOk = True
    
    Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    IsOk = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
