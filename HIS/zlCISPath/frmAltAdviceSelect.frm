VERSION 5.00
Begin VB.Form frmAltAdviceSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ѡҽ��"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11730
   Icon            =   "frmAltAdviceSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -120
      TabIndex        =   3
      Top             =   5760
      Width           =   11895
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10440
      TabIndex        =   2
      Top             =   5925
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   9120
      TabIndex        =   1
      Top             =   5925
      Width           =   1100
   End
   Begin zlCISPath.UCAdviceList UCAdvice 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   -50
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
   End
End
Attribute VB_Name = "frmAltAdviceSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng·����ĿID As Long
Private mstrSelectedIds As String
Private mblnOK As Boolean
Private mrsAdvice As Recordset
Private mintӤ�� As Integer
Private mintFunc As Integer '0-ȱʡΪסԺ�ٴ�·��;1-�����ٴ�·��

Public Function ShowSelect(ByVal frmParent As Object, ByVal lng·����ĿID As Long, Optional ByVal strSelectedIDs As String, _
    Optional ByVal intӤ�� As Integer, Optional ByVal intFunc As Integer) As String
'���ܣ�����ѡ����棬����ѡ����ҽ��IDs
    mlng·����ĿID = lng·����ĿID
    mstrSelectedIds = strSelectedIDs
    mintӤ�� = intӤ��
    mintFunc = intFunc
    Me.Show 1, frmParent
    ShowSelect = IIf(mblnOK, mstrSelectedIds, strSelectedIDs)
End Function

Private Sub ShowAltAdvice()
'���ܣ���ʾ��ѡҽ��
    Dim strSQL As String, rstmp As Recordset
    Dim i As Long
    
    On Error GoTo errH
    If mintFunc = 0 Then
        strSql = _
            "Select Distinct (a.Id * 10 +" & mintӤ�� & ") as id, Decode(a.���ID,NULL,NULL,(a.���id * 10 + " & mintӤ�� & ")) AS ���id, a.���, a.��Ч, a.������Ŀid, a.�շ�ϸĿid, a.ҽ������, a.��������, a.�ܸ�����, a.�걾��λ, a.��鷽��, a.ҽ������, a.ִ��Ƶ��," & vbNewLine & _
                    "                a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, a.ִ������,a.ִ�б��,a.�����ĿID, a.ִ�п���id, a.ʱ�䷽��, a.�Ƿ�ȱʡ, Decode(instr(',' || [4] || ',',',' || (a.ID *10 + " & mintӤ�� & ") || ','), 0, 0, 1) As �Ƿ�ѡ" & vbNewLine & _
                    "From ·��ҽ������ A, �ٴ�·��ҽ�� B, �ٴ�·����Ŀ C" & vbNewLine & _
                    "Where a.Id = b.ҽ������id And b.·����Ŀid = c.Id And c.id=[3]" & vbNewLine & _
                    "Order By a.���, a.Id"
    Else
        strSql = _
            "Select Distinct (a.Id * 10 +" & mintӤ�� & ") as id, Decode(a.���ID,NULL,NULL,(a.���id * 10 + " & mintӤ�� & ")) AS ���id, a.���, a.��Ч, a.������Ŀid, a.�շ�ϸĿid, a.ҽ������, a.��������, a.�ܸ�����, a.�걾��λ, a.��鷽��, a.ҽ������, a.ִ��Ƶ��," & vbNewLine & _
                    "                a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, a.ִ������,a.ִ�б��,a.�����ĿID, a.ִ�п���id, a.ʱ�䷽��, a.�Ƿ�ȱʡ, Decode(instr(',' || [4] || ',',',' || (a.ID *10 + " & mintӤ�� & ") || ','), 0, 0, 1) As �Ƿ�ѡ" & vbNewLine & _
                    "From ����·��ҽ������ A, ����·��ҽ�� B, ����·����Ŀ C" & vbNewLine & _
                    "Where a.Id = b.ҽ������id And b.·����Ŀid = c.Id And c.id=[3]" & vbNewLine & _
                    "Order By a.���, a.Id"
    End If
    Call UCAdvice.ShowAdvice(3, strSql, 0, 0, , mlng·����ĿID, mstrSelectedIds)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrSelectedIds = UCAdvice.GetAdviceIDSelected(1)
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Call ShowAltAdvice
End Sub
