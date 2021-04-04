VERSION 5.00
Begin VB.Form frmParameter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmParameter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4560
   StartUpPosition =   1  '����������
   Begin VB.Frame fraPriceFolw 
      Caption         =   "��������"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4335
      Begin VB.CheckBox chkPriceFlow 
         Caption         =   "������Ҫ���"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   0
      Top             =   1920
      Width           =   1100
   End
End
Attribute VB_Name = "frmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnLoad As Boolean '�����Ƿ������� true-��� false-δ���

Public Sub ShowMe(ByVal fraParent As Form)
    Me.Show vbModal, fraParent
End Sub

Private Sub LoadData()
    Dim int���� As Integer
    
    int���� = zldatabase.GetPara("������Ҫ���", glngSys, 1009, 0)
    chkPriceFlow.Value = IIF(int���� = 1, 1, 0)
End Sub


Private Sub chkPriceFlow_Click()
    Dim blnResult As Boolean
    
    If mblnLoad = True Then
        blnResult = checkNotPrice
        If blnResult = False Then
            MsgBox "������δ��Ч�ĵ��۵��ݣ������޸Ĵ˲�����", vbInformation, gstrSysName
            chkPriceFlow.Value = 1
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnLoad = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    zldatabase.SetPara "������Ҫ���", IIF(chkPriceFlow.Value = 1, "1", "0"), glngSys, 1009
    Unload Me
End Sub

Private Sub Form_Load()
    Call LoadData
    mblnLoad = True
End Sub

Private Function checkNotPrice() As Boolean
    '����Ƿ񻹴���δ��Ч�ļ۸�
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If chkPriceFlow.Value = 0 Then
        gstrSQL = "Select 1 From �շѵ��ۼ�¼ Where ��˱�־ = 0 And Rownum <= 1"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "δ��Ч���ݲ�ѯ")
        If rsData.RecordCount > 0 Then
            checkNotPrice = False
        Else
            checkNotPrice = True
        End If
    Else
        checkNotPrice = True
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
