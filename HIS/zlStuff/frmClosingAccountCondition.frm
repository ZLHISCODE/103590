VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClosingAccountCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ֹ��������"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4440
   Icon            =   "frmClosingAccountCondition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3235
      TabIndex        =   2
      Top             =   1800
      Width           =   1100
   End
   Begin VB.Frame fraConditiom 
      Caption         =   "��ĩ����ѡ��"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton optָ��ʱ�� 
         Caption         =   "ָ��ʱ��"
         Height          =   180
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton opt��ǰʱ�� 
         Caption         =   "��ǰʱ��"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   900
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   194510851
         CurrentDate     =   36901
      End
   End
End
Attribute VB_Name = "frmClosingAccountCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng�ⷿID As Long
Private mblnSelect As Boolean
Private mstr���ʱ�� As String

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Public Function GetCondition(frmMain As Form, ByVal lng�ⷿID As Long, ByRef str���ʱ��) As Boolean
    'ѡ��ǰʱ�䣬����str���ʱ��=""��ѡ��ָ��ʱ�䣬����str���ʱ��Ϊ����ʱ�䣻
    'GetCondition��true-��棻false-ȡ�����
    mlng�ⷿID = lng�ⷿID
    mblnSelect = False
    
    Me.Show 1, frmMain
    
    str���ʱ�� = mstr���ʱ��
    GetCondition = mblnSelect
    
End Function

Private Sub CmdSave_Click()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    mstr���ʱ�� = IIf(optָ��ʱ��.Value = True, Format(dtpDate.Value, "yyyy-MM-dd hh:mm:ss"), "")
    
    If optָ��ʱ��.Value = True Then 'ָ��ʱ��Ҫ���ʱ���Ƿ����������ĩʱ��
        gstrSQL = " Select Max(��ĩ����) ����ĩ����, Max(��ĩ����) + 1 / 24 / 60 / 60 �ڳ����� From ���Ͻ���¼ Where �ⷿid = [1] And ȡ���� Is Null "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", mlng�ⷿID)
        
        If rsTemp.EOF = True Or IsNull(rsTemp!����ĩ����) = True Then
            MsgBox "�ÿⷿû�н���¼�����ȳ�ʼ����", vbInformation, gstrSysName
            mblnSelect = False
        Else
            If mstr���ʱ�� < Format(rsTemp!�ڳ�����, "yyyy-MM-dd hh:mm:ss") Then
                MsgBox "ָ��ʱ��������������ĩ���ڣ�" & Format(rsTemp!����ĩ����, "yyyy-MM-dd hh:mm:ss") & "����", vbInformation, gstrSysName
                dtpDate.SetFocus
                Exit Sub
            End If
            
            If mstr���ʱ�� > Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") Then
                MsgBox "ָ��ʱ�䲻�ܴ��ڵ�ǰϵͳʱ�䣡", vbInformation, gstrSysName
                dtpDate.SetFocus
                Exit Sub
            End If
            
            mblnSelect = True
        End If
    Else 'ѡ��ǰʱ�䲻��У��
        mblnSelect = True
    End If
    
    Unload Me
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()

    dtpDate.Value = Format(zlDatabase.Currentdate, dtpDate.CustomFormat)
    dtpDate.Enabled = optָ��ʱ��.Value = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub opt��ǰʱ��_Click()
    dtpDate.Enabled = optָ��ʱ��.Value = True
End Sub

Private Sub optָ��ʱ��_Click()
    dtpDate.Enabled = optָ��ʱ��.Value = True
End Sub
