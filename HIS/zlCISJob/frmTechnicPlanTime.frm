VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTechnicPlanTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ʱ�䰲��"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   3735
   Icon            =   "frmTechnicPlanTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2220
      TabIndex        =   2
      Top             =   1785
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   0
      TabIndex        =   3
      Top             =   1590
      Width           =   4440
   End
   Begin MSComCtl2.DTPicker dtpPlan 
      Height          =   300
      Left            =   1530
      TabIndex        =   0
      Top             =   1035
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   118161411
      CurrentDate     =   39158
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1125
      TabIndex        =   1
      Top             =   1785
      Width           =   1100
   End
   Begin VB.Label lblInfo 
      Caption         =   "����Ŀ��Ҫ��ִ��ʱ��Ϊ��yyyy-MM-dd HH:mm�������Ը���ʵ�ʹ�����������¶�ʱ����а���"
      Height          =   705
      Left            =   990
      TabIndex        =   5
      Top             =   165
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ��ʱ��"
      Height          =   180
      Left            =   720
      TabIndex        =   4
      Top             =   1095
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   135
      Picture         =   "frmTechnicPlanTime.frx":058A
      Top             =   195
      Width           =   720
   End
End
Attribute VB_Name = "frmTechnicPlanTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mlngҽ��ID As Long
Private mlng���ͺ� As Long
Private mlngִ�п���ID As Long
Private mdҪ��ʱ�� As Date
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mrs������Ϣ As ADODB.Recordset

Public Function ShowMe(frmParent As Object, ByRef objMip As Object, ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, Optional ByVal lngִ�п���ID As Long) As Boolean
    mlngҽ��ID = lngҽ��ID
    mlng���ͺ� = lng���ͺ�
    mlngִ�п���ID = lngִ�п���ID
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Me.Show 1, frmParent
    
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim StrSQL As String
    Dim blnTrans As Boolean
    
    If Format(dtpPlan.Value, "yyyy-MM-dd HH:mm") < Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") Then
        MsgBox "���°��ŵ�ִ��ʱ��Ӧ���ڵ�ǰʱ��֮��", vbInformation, gstrSysName
        dtpPlan.SetFocus: Exit Sub
    End If
    If Format(dtpPlan.Value, "yyyy-MM-dd HH:mm") = Format(mdҪ��ʱ��, "yyyy-MM-dd HH:mm") Then
        If MsgBox("��ǰ���ŵ�ִ��ʱ����ԭ��Ҫ���ʱ����ͬ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            dtpPlan.SetFocus: Exit Sub
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    If mlngִ�п���ID <> 0 Then
        StrSQL = "Zl_����ҽ������_���ұ��(" & mlngҽ��ID & "," & mlng���ͺ� & "," & mlngִ�п���ID & ")"
        Call zlDatabase.ExecuteProcedure(StrSQL, Me.Caption)
    End If
    StrSQL = "zl_����ҽ��ִ��_Arrange(" & mlngҽ��ID & "," & mlng���ͺ� & ",To_Date('" & Format(dtpPlan.Value, "yyyy-MM-dd HH:mm:00") & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(StrSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
  
    With mrs������Ϣ
        Call ZLHIS_CIS_005(mclsMipModule, Val(!����ID & ""), !���� & "", !סԺ�� & "", , 2, Val(!��ҳID & ""), Val(!��ǰ����ID & ""), Val(!��ǰ����id & ""), "", , !��ǰ���� & "", _
            mlngҽ��ID, Val(!ҽ����Ч & ""), !������� & "", !�������� & "", Val(!������Ŀid & ""), !ҽ������ & "", Format(dtpPlan.Value, "yyyy-MM-dd HH:mm:00"), "")
    End With
            
    mblnOk = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Load()
    Dim StrSQL As String
    
    mblnOk = False
    
    On Error GoTo errH
    
    StrSQL = "Select Sysdate As ��ǰʱ��, b.����ʱ��, Decode(Nvl(a.ҽ����Ч, 0), 1, a.��ʼִ��ʱ��, b.�״�ʱ��) As Ҫ��ʱ��, a.����id, a.����, c.סԺ��, a.��ҳid," & vbNewLine & _
        "       c.��ǰ����id, c.��ǰ����id, c.��ǰ����, a.ҽ����Ч, a.������Ŀid, a.ҽ������, a.�������, d.��������" & vbNewLine & _
        "From ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C, ������ĿĿ¼ D" & vbNewLine & _
        "Where a.Id = b.ҽ��id And a.����id = c.����id And a.������Ŀid = d.Id And b.ҽ��id = [1] And b.���ͺ� = [2]"
    Set mrs������Ϣ = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�)
    mdҪ��ʱ�� = mrs������Ϣ!Ҫ��ʱ��
    lblInfo.Caption = Replace(lblInfo.Caption, "yyyy-MM-dd HH:mm", Format(mdҪ��ʱ��, "yyyy-MM-dd HH:mm"))
    If Not IsNull(mrs������Ϣ!����ʱ��) Then
        dtpPlan.Value = Format(mrs������Ϣ!����ʱ��, "yyyy-MM-dd HH:mm")
    Else
        dtpPlan.Value = Format(mrs������Ϣ!Ҫ��ʱ��, "yyyy-MM-dd HH:mm")
    End If
    If Format(dtpPlan.Value, "yyyy-MM-dd HH:mm") < Format(mrs������Ϣ!��ǰʱ��, "yyyy-MM-dd HH:mm") Then
        dtpPlan.Value = DateAdd("n", 30, mrs������Ϣ!��ǰʱ��)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrs������Ϣ = Nothing
    Set mclsMipModule = Nothing
End Sub
