VERSION 5.00
Begin VB.Form frmDiseaseReportIncept 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "疾病报告接收"
   ClientHeight    =   3900
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5655
   Icon            =   "frmDiseaseReportIncept.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3195
      TabIndex        =   10
      Top             =   3360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4305
      TabIndex        =   11
      Top             =   3360
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -30
      TabIndex        =   9
      Top             =   3195
      Width           =   6030
   End
   Begin VB.TextBox txtComment 
      Height          =   660
      Left            =   795
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2370
      Width           =   4605
   End
   Begin VB.OptionButton optIncept 
      Caption         =   "拒绝接收(&R)"
      Height          =   225
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   1305
   End
   Begin VB.OptionButton optIncept 
      Caption         =   "同意接收(&A)"
      Height          =   225
      Index           =   0
      Left            =   780
      TabIndex        =   5
      Top             =   1800
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -15
      TabIndex        =   0
      Top             =   600
      Width           =   6030
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "同意或拒绝的理由(&M):"
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   2130
      Width           =   1800
   End
   Begin VB.Label lblWriter 
      AutoSize        =   -1  'True
      Caption         =   "填报人:"
      Height          =   180
      Left            =   780
      TabIndex        =   4
      Top             =   1350
      Width           =   630
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "病人:"
      Height          =   180
      Left            =   780
      TabIndex        =   3
      Top             =   1065
      Width           =   450
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "报告:"
      Height          =   180
      Left            =   780
      TabIndex        =   2
      Top             =   780
      Width           =   450
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   150
      Picture         =   "frmDiseaseReportIncept.frx":038A
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "仔细检查临床医生填写的疾病报告内容是否符合要求，决定接收或拒绝该疾病报告。"
      Height          =   360
      Left            =   780
      TabIndex        =   1
      Top             =   135
      Width           =   4680
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDiseaseReportIncept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOk As Boolean
Private mstrId As String
Private mlngPatiId As Long, mlngPageId As Long, mlngFrom As Long

Public Function ShowMe(ByVal frmParent As Object, strRecordId As String, strInfo As String) As Boolean
'strInfo=报告|科室|姓名|性别|年龄|就诊号|填报人|填报时间|病人ID|主页ID
'新病历 strInfo=strInfo & |甲类传染病|乙类传染病|丙类传染病|病例分类|病例分类2
    mstrId = strRecordId
    
    Err = 0: On Error GoTo errHand

    Me.lblFile.Caption = "报告:" & Split(strInfo, "|")(0) & "    科室:" & Split(strInfo, "|")(1)
    Me.lblPati.Caption = "病人: " & Split(strInfo, "|")(2) & "," & Split(strInfo, "|")(3) & "," & Split(strInfo, "|")(4) & "  " & Split(strInfo, "|")(5)
    Me.lblWriter.Caption = "填报人:" & Split(strInfo, "|")(6) & "    填报时间:" & Split(strInfo, "|")(7)
    Me.lblFile.Tag = Split(strInfo, "|")(1)
    mlngPatiId = Split(strInfo, "|")(8): mlngPageId = Split(strInfo, "|")(9)
    If InStr(Split(strInfo, "|")(1), "门诊:") > 0 Then
        mlngFrom = 1
    ElseIf InStr(Split(strInfo, "|")(1), "住院:") > 0 Then
        mlngFrom = 2
    End If
    
    If Not IsNumeric(mstrId) Then
        lblComment.Tag = Split(strInfo, "|")(10) & ";" & Split(strInfo, "|")(11) & ";" & Split(strInfo, "|")(12)
        txtComment.Tag = Split(strInfo, "|")(13) & ";" & Split(strInfo, "|")(14)
    End If
    
    Me.Show vbModal, frmParent
    
    ShowMe = mblnOk
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    mblnOk = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    Dim lngKey As Long, rsTemp As ADODB.Recordset, strContent As String, dblDocID As Double
    
    If LenB(StrConv(Trim(Me.txtComment.Text), vbFromUnicode)) > Me.txtComment.MaxLength Then
        MsgBox "说明超长（最多" & Me.txtComment.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName
        Me.txtComment.SetFocus: Exit Sub
    End If
    
    If IsNumeric(mstrId) Then
        strContent = ""
        gstrSQL = "Zl_疾病申报记录_Incept(" & CDbl(mstrId) & "," & IIf(Me.optIncept(0).Value, 1, -1) & ",'" & Trim(Me.txtComment.Text) & "','" & mstrId & "'," & mlngPatiId & "," & mlngPageId & "," & mlngFrom & ",'')"
    Else
        '新版病历，将GUID转数据，以确保 疾病申报记录PK
        gstrSQL = "Select 文件ID From 疾病申报记录 Where 文档ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", mstrId)
        If rsTemp.EOF Then
            dblDocID = gobjEmr.GuidToHashCode(mstrId)
        Else
            dblDocID = rsTemp!文件ID
        End If
        '新版病历，获取申报信息
        strContent = lblComment.Tag & "|" & txtComment.Tag
        gstrSQL = "Zl_疾病申报记录_Incept(" & dblDocID & "," & IIf(Me.optIncept(0).Value, 1, -1) & ",'" & Trim(Me.txtComment.Text) & "','" & mstrId & "'," & mlngPatiId & "," & mlngPageId & "," & mlngFrom & ",'" & strContent & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If InStr(lblFile.Tag, "门诊") = 0 Then
        If optIncept(1).Value = True Then  '当拒绝时，当做一次抽查结果记录
            lngKey = zlDatabase.GetNextId("病案反馈记录")
            gstrSQL = "zl_病案反馈记录_Update(" & lngKey & ",Null,Null," & mlngPatiId & "," & mlngPageId & ",7,'" & mstrId & "','" & _
                    txtComment.Text & "',Null,'" & gstrUserName & "',To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')" & _
                    ",To_Date('" & Format(zlDatabase.Currentdate + 1, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Else '接收时,对自已曾经拒绝过的记录做完成标志
            gstrSQL = "Select ID From 病案反馈记录 Where 病人ID=[1] And 主页ID=[2] and 反馈对象=7 And 文件ID=[3] And 反馈人=[4]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mstrId, gstrUserName)
            Do Until rsTemp.EOF
                gstrSQL = "zl_病案反馈记录_Finish(" & rsTemp!ID & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                rsTemp.MoveNext
            Loop
        End If
    End If
    
    mblnOk = True: Me.Hide: Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optIncept_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtComment_GotFocus()
    Me.txtComment.SelStart = 0: Me.txtComment.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtComment_LostFocus()
    Me.txtComment.Text = Replace(Me.txtComment, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub



