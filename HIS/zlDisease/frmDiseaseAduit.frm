VERSION 5.00
Begin VB.Form frmDiseaseAduit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "传染病报告审核"
   ClientHeight    =   3810
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6135
   Icon            =   "frmDiseaseAduit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3225
      TabIndex        =   6
      Top             =   3285
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4335
      TabIndex        =   5
      Top             =   3285
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   6030
   End
   Begin VB.TextBox txtContent 
      Height          =   660
      Left            =   825
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2295
      Width           =   4605
   End
   Begin VB.OptionButton optAduit 
      Caption         =   "要求返修"
      Height          =   225
      Index           =   1
      Left            =   2430
      TabIndex        =   2
      Top             =   1725
      Width           =   1305
   End
   Begin VB.OptionButton optAduit 
      Caption         =   "审核通过"
      Height          =   225
      Index           =   0
      Left            =   810
      TabIndex        =   1
      Top             =   1725
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   645
      Width           =   6030
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "通过或返修的理由(&M):"
      Height          =   180
      Left            =   810
      TabIndex        =   11
      Top             =   2055
      Width           =   1800
   End
   Begin VB.Label lblWriter 
      AutoSize        =   -1  'True
      Caption         =   "填报人:"
      Height          =   180
      Left            =   810
      TabIndex        =   10
      Top             =   1275
      Width           =   630
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "病人:"
      Height          =   180
      Left            =   810
      TabIndex        =   9
      Top             =   990
      Width           =   450
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "报告:"
      Height          =   180
      Left            =   810
      TabIndex        =   8
      Top             =   705
      Width           =   450
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   180
      Picture         =   "frmDiseaseAduit.frx":6852
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "仔细检查临床医生填写的疾病报告内容是否符合要求，决定通过或要求返修该疾病报告。"
      Height          =   360
      Left            =   810
      TabIndex        =   7
      Top             =   180
      Width           =   4680
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDiseaseAduit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrReportID As String
Private mintResult As Integer
Private mblnOk As Boolean
Private mlngPatiId As Long
Private mlngPageId As Long
Private mlngFrom As Long

Public Function ShowDiseaseAudit(ByVal frmParent As Object, ByVal strID As String, ByVal strInfo As String, ByRef intAduitState As Integer) As Boolean
    On Error GoTo errHand
    mstrReportID = strID
    mlngPatiId = Split(strInfo, "|")(5)
    mlngPageId = Split(strInfo, "|")(6)
    mlngFrom = Val(Split(strInfo, "|")(7))
    
    Me.lblFile.Caption = "报告:" & Split(strInfo, "|")(0) & "    科室:" & Split(strInfo, "|")(1)
    Me.lblPati.Caption = "病人: " & Split(strInfo, "|")(2) & "," & Split(strInfo, "|")(3) & "," & Split(strInfo, "|")(4)
    Me.lblWriter.Caption = "填报人:" & Split(strInfo, "|")(8) & "    填报时间:" & Split(strInfo, "|")(9)

    Me.Show 1, frmParent
    ShowDiseaseAudit = mblnOk
    intAduitState = mintResult
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ProcessReport(ByVal blnResult As Boolean) As Boolean
'功能：对报告进行接收审核
'参数:blnResult - true 审核通过；false-要求返修
    On Error GoTo errH:
    Dim strContent As String
    Dim strSQLIncept As String, strSQLAduit As String
    Dim blnTrans As Boolean
    Dim str审核时间 As String, str审核医生 As String, str审核说明 As String
    
    If LenB(StrConv(Trim(Me.txtContent.Text), vbFromUnicode)) > Me.txtContent.MaxLength Then
        MsgBox "说明超长（最多" & Me.txtContent.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName
        Me.txtContent.SetFocus: Exit Function
    End If
    
    str审核时间 = "to_date('" & Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    str审核医生 = "'" & UserInfo.姓名 & "'"
    str审核说明 = "'" & Trim(txtContent.Text) & "'"
    mintResult = IIf(blnResult, 3, 4)

    If IsNumeric(mstrReportID) Then
        strContent = ""
        strSQLIncept = "Zl_疾病申报记录_Incept(" & CDbl(mstrReportID) & "," & 1 & ",NULL,'" & mstrReportID & "'," & mlngPatiId & "," & mlngPageId & "," & mlngFrom & ",'')"
        strSQLAduit = "Zl_疾病申报记录_Update(" & CDbl(mstrReportID) & ", " & CStr(mintResult) & "," & str审核时间 & "," & str审核医生 & "," & str审核说明 & ",NULL,NULL,NULL )"
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQLIncept, Me.Caption)
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQLAduit, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    
    mintResult = IIf(blnResult, 3, 4)
    ProcessReport = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    If LenB(StrConv(Trim(Me.txtContent.Text), vbFromUnicode)) > Me.txtContent.MaxLength Then
        MsgBox "说明超长（最多" & Me.txtContent.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName
        Me.txtContent.SetFocus: Exit Sub
    End If
    On Error GoTo errH
    strSql = "select t.最后版本 from 电子病历记录 t where t.id =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "查询最新文件处理状态", mstrReportID)
    If rsTmp.RecordCount <> 0 Then
        If Val(rsTmp!最后版本 & "") = 0 Then
            MsgBox "当前文件状态已被改变不能进行审核,请刷新后再试!", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    
    mblnOk = True

    Call ProcessReport(optAduit(0).Value)
    Unload Me
	Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optAduit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call gobjComlib.ZLCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtContent_GotFocus()
    Me.txtContent.SelStart = 0: Me.txtContent.SelLength = 1000
    Call gobjComlib.ZLCommFun.OpenIme(True)
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call gobjComlib.ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtContent_LostFocus()
    Me.txtContent.Text = Replace(Me.txtContent, Chr(vbKeyReturn), "")
    Call gobjComlib.ZLCommFun.OpenIme(False)
End Sub

