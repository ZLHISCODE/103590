VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTechnicPlanTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "时间安排"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
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
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1125
      TabIndex        =   1
      Top             =   1785
      Width           =   1100
   End
   Begin VB.Label lblInfo 
      Caption         =   "该项目的要求执行时间为：yyyy-MM-dd HH:mm，您可以根据实际工作的情况重新对时间进行安排"
      Height          =   705
      Left            =   990
      TabIndex        =   5
      Top             =   165
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行时间"
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
Private mlng医嘱ID As Long
Private mlng发送号 As Long
Private mlng执行科室ID As Long
Private md要求时间 As Date
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mrs基本信息 As ADODB.Recordset

Public Function ShowMe(frmParent As Object, ByRef objMip As Object, ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, Optional ByVal lng执行科室ID As Long) As Boolean
    mlng医嘱ID = lng医嘱ID
    mlng发送号 = lng发送号
    mlng执行科室ID = lng执行科室ID
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
        MsgBox "重新安排的执行时间应该在当前时间之后。", vbInformation, gstrSysName
        dtpPlan.SetFocus: Exit Sub
    End If
    If Format(dtpPlan.Value, "yyyy-MM-dd HH:mm") = Format(md要求时间, "yyyy-MM-dd HH:mm") Then
        If MsgBox("当前安排的执行时间与原本要求的时间相同，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            dtpPlan.SetFocus: Exit Sub
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    If mlng执行科室ID <> 0 Then
        StrSQL = "Zl_病人医嘱发送_科室变更(" & mlng医嘱ID & "," & mlng发送号 & "," & mlng执行科室ID & ")"
        Call zlDatabase.ExecuteProcedure(StrSQL, Me.Caption)
    End If
    StrSQL = "zl_病人医嘱执行_Arrange(" & mlng医嘱ID & "," & mlng发送号 & ",To_Date('" & Format(dtpPlan.Value, "yyyy-MM-dd HH:mm:00") & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(StrSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
  
    With mrs基本信息
        Call ZLHIS_CIS_005(mclsMipModule, Val(!病人ID & ""), !姓名 & "", !住院号 & "", , 2, Val(!主页ID & ""), Val(!当前病区ID & ""), Val(!当前科室id & ""), "", , !当前床号 & "", _
            mlng医嘱ID, Val(!医嘱期效 & ""), !诊疗类别 & "", !操作类型 & "", Val(!诊疗项目id & ""), !医嘱内容 & "", Format(dtpPlan.Value, "yyyy-MM-dd HH:mm:00"), "")
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
    
    StrSQL = "Select Sysdate As 当前时间, b.安排时间, Decode(Nvl(a.医嘱期效, 0), 1, a.开始执行时间, b.首次时间) As 要求时间, a.病人id, a.姓名, c.住院号, a.主页id," & vbNewLine & _
        "       c.当前科室id, c.当前病区id, c.当前床号, a.医嘱期效, a.诊疗项目id, a.医嘱内容, a.诊疗类别, d.操作类型" & vbNewLine & _
        "From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C, 诊疗项目目录 D" & vbNewLine & _
        "Where a.Id = b.医嘱id And a.病人id = c.病人id And a.诊疗项目id = d.Id And b.医嘱id = [1] And b.发送号 = [2]"
    Set mrs基本信息 = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng医嘱ID, mlng发送号)
    md要求时间 = mrs基本信息!要求时间
    lblInfo.Caption = Replace(lblInfo.Caption, "yyyy-MM-dd HH:mm", Format(md要求时间, "yyyy-MM-dd HH:mm"))
    If Not IsNull(mrs基本信息!安排时间) Then
        dtpPlan.Value = Format(mrs基本信息!安排时间, "yyyy-MM-dd HH:mm")
    Else
        dtpPlan.Value = Format(mrs基本信息!要求时间, "yyyy-MM-dd HH:mm")
    End If
    If Format(dtpPlan.Value, "yyyy-MM-dd HH:mm") < Format(mrs基本信息!当前时间, "yyyy-MM-dd HH:mm") Then
        dtpPlan.Value = DateAdd("n", 30, mrs基本信息!当前时间)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrs基本信息 = Nothing
    Set mclsMipModule = Nothing
End Sub
