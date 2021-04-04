VERSION 5.00
Begin VB.Form frmEInvoiceParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   Icon            =   "frmEInvoiceParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4590
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd��ӡ���� 
      Caption         =   "��֪����ӡ����(&P)"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   2130
      Width           =   2370
   End
   Begin VB.Frame fra 
      Caption         =   "��Ʊ����뷽ʽ"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2385
      Begin VB.OptionButton Option���뷽ʽ 
         Caption         =   "���ͻ��˺��շ�Ա"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option���뷽ʽ 
         Caption         =   "���շ�Ա"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   1095
      End
      Begin VB.OptionButton Option���뷽ʽ 
         Caption         =   "���ͻ���"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3090
      TabIndex        =   5
      Top             =   330
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3090
      TabIndex        =   6
      Top             =   810
      Width           =   1100
   End
End
Attribute VB_Name = "frmEInvoiceParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln�����Ʊ����� As Boolean
Private mlngSys As Long
Private mlngModule As Long
Private mstrPrivs As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intTmp As Integer, blnSetUp As Boolean
    Dim strSQL As String
    
    blnSetUp = InStr(1, mstrPrivs, ";��������;") > 0
    intTmp = IIf(Option���뷽ʽ(2).Value, 2, IIf(Option���뷽ʽ(1).Value, 1, 0))
    If fra.Tag <> intTmp Then
        zlDatabase.SetPara "��Ʊ����뷽ʽ", intTmp, mlngSys, mlngModule, blnSetUp
        If mbln�����Ʊ����� Then
            strSQL = "Zl_Ʊ�ݿ�Ʊ�����_Update(3)"
            Call zlDatabase.ExecuteProcedure(strSQL, "Ʊ�ݿ�Ʊ�����")
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitPara
End Sub

Private Sub InitPara()
    '��ʼ������
    Dim intTmp As Integer, blnSetUp As Boolean
   
    mstrPrivs = ";" & GetPrivFunc(mlngSys, mlngModule) & ";"
    blnSetUp = InStr(1, mstrPrivs, ";��������;") > 0
    
    intTmp = zlDatabase.GetPara("��Ʊ����뷽ʽ", mlngSys, mlngModule, 1, Option���뷽ʽ, blnSetUp)
    fra.Tag = intTmp
    Option���뷽ʽ(intTmp).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln�����Ʊ����� = False
End Sub

Private Sub Option���뷽ʽ_Click(Index As Integer)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    If Val(fra.Tag) = Index Then Exit Sub
    strSQL = "select  1 from Ʊ�ݿ�Ʊ����� Where Rownum<2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.EOF Then Exit Sub
    If MsgBox("��ȷ��Ҫ���ġ���Ʊ����뷽ʽ���𣬸����˱��������������Ʊ�ݿ�Ʊ����ա����е����ݣ�", _
       vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        mbln�����Ʊ����� = True
    Else
        Option���뷽ʽ(Val(fra.Tag)).Value = True
        zlControl.ControlSetFocus Option���뷽ʽ(Val(fra.Tag))
    End If
End Sub

Public Sub ShowMe(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long)
    On Error GoTo errHandle
    mlngSys = lngSys: mlngModule = lngModule
    Me.Show 1, frmMain
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


