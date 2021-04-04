VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNOSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����סԺ�Ų�ѯ"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmNOSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ListBox lst 
      Height          =   3300
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   1290
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   5
      Top             =   3625
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   3120
      Width           =   1290
   End
   Begin MSComCtl2.DTPicker dtp��ԺB 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   93847555
      CurrentDate     =   37003
   End
   Begin MSComCtl2.DTPicker dtp��ԺE 
      Height          =   300
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   93847555
      CurrentDate     =   37003
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ժʱ��"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   2340
      TabIndex        =   6
      Top             =   180
      Width           =   180
   End
End
Attribute VB_Name = "frmNOSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrNO As String
Private mDatBegin As Date
Private mDatEnd As Date


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lst.ListIndex = -1 Then Exit Sub
    
    mstrNO = lst.List(lst.ListIndex)
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim DatBegin As Date, DatEnd As Date, i As Long
    Dim strMinNO As String, strMaxNO As String  'ʹ���ַ���,��Ϊ����16λ�������ͻᱻת��Ϊ��ѧ������
    
    If dtp��ԺB.Value > dtp��ԺE.Value Then
        MsgBox "��ʼʱ�䲻�ܴ��ڽ���ʱ��!", vbInformation, gstrSysName
        Exit Sub
    Else
        DatBegin = dtp��ԺB.Value
        DatEnd = dtp��ԺE.Value
    End If
    lst.Clear
 
    strSQL = "Select Min(סԺ��) MinסԺ��, Max(סԺ��) MaxסԺ��" & vbNewLine & _
            "From ������ҳ" & vbNewLine & _
            "Where ��Ժ���� Between [1] And [2] And סԺ�� Is Not Null"
    On Error GoTo errH
    
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, App.ProductName, DatBegin, DatEnd)
    If rsTmp.RecordCount = 0 Then Exit Sub
    If IsNull(rsTmp!MinסԺ��) Or IsNull(rsTmp!MaxסԺ��) Then Exit Sub
    
    If rsTmp!MinסԺ�� - 10 < 10 Then   '���ܵ�ǰʱ�䷶Χ�ڵ���СסԺ��֮ǰȡ����һЩסԺ��,�������ȡ����10��
        strMinNO = 1
    Else
        strMinNO = rsTmp!MinסԺ�� - 10
    End If
    strMaxNO = rsTmp!MaxסԺ�� - 1
    If strMaxNO - strMinNO > 10000 Then strMaxNO = strMinNO + 10000
            
    strSQL = "Select סԺ��" & vbNewLine & _
            "From (Select Rownum + ([3] - 1) סԺ��" & vbNewLine & _
            "       From Dual A" & vbNewLine & _
            "       Connect By Rownum <= [4] - ([3] - 1)" & vbNewLine & _
            "       Minus" & vbNewLine & _
            "       Select סԺ�� From ������ҳ Where ��Ժ���� Between [1] And [2] And סԺ�� Is Not Null) A" & vbNewLine & _
            "Where Not Exists (Select 1 From ������ҳ B Where B.סԺ�� = A.סԺ��)" & vbNewLine & _
            "Order By סԺ��"
    '������ЩסԺ�������ù���,ֻ�����ڵ�ǰʱ�䷶Χ��û�дӲ�����ҳ���ҵ�,��Ϊ������ʹ�ô��סԺ��,����Ҫ��Exists�ж�
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, App.ProductName, DatBegin, DatEnd, strMinNO, strMaxNO)

    For i = 1 To rsTmp.RecordCount
        lst.AddItem rsTmp!סԺ��
        rsTmp.MoveNext
    Next
    If lst.ListCount > 0 Then lst.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtp��ԺE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub Form_Activate()
    If lst.ListCount > 0 Then lst.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is dtp��ԺE Or Me.ActiveControl Is dtp��ԺB Then
             SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    
    If mDatBegin = CDate(0) Or mDatEnd = CDate(0) Then
        mDatEnd = zldatabase.Currentdate
        mDatBegin = DateAdd("d", -7, mDatEnd)
    End If
    
    dtp��ԺB.Value = Format(mDatBegin, "yyyy-MM-dd 00:00:00")
    dtp��ԺE.Value = Format(mDatEnd, "yyyy-MM-dd 23:59:59")
    
    Call cmdRefresh_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mDatBegin = dtp��ԺB.Value
    mDatEnd = dtp��ԺE.Value
End Sub

Private Sub lst_DblClick()
    Call cmdOK_Click
End Sub


Public Sub ShowMe(objParent As Object, ByRef strno As String)
'����:
'����:ѡ���סԺ��
    Call Me.Show(vbModal, objParent)
    strno = mstrNO
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lst.ListIndex <> -1 Then
            Call cmdOK_Click
        End If
    End If
End Sub
