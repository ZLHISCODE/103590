VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholAntibody_FeedbackUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "抗体反馈"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   Icon            =   "frmPatholAntibody_FeedbackUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtShow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3135
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.ComboBox cbxExperimentType 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1425
      Width           =   2055
   End
   Begin VB.TextBox txtReferencePatholNum 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   915
      Width           =   4695
   End
   Begin VB.CommandButton cmdFeedback_Sure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   3480
      TabIndex        =   6
      Top             =   3915
      Width           =   1215
   End
   Begin VB.CommandButton cmdFeedback_Cancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   4800
      TabIndex        =   7
      Top             =   3915
      Width           =   1215
   End
   Begin VB.TextBox txtFeedbackAdvice 
      Height          =   780
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2955
      Width           =   4425
   End
   Begin VB.ComboBox cbxAntibodyGrade 
      Height          =   300
      Left            =   4800
      TabIndex        =   2
      Text            =   "良"
      Top             =   1425
      Width           =   1210
   End
   Begin VB.TextBox txtFeedbackDoctor 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   2445
      Width           =   4665
   End
   Begin MSComCtl2.DTPicker dtpFeedbackTime 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1935
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd hh:mm"
      Format          =   155058179
      CurrentDate     =   40646.4399652778
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  反馈时间："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   1980
      Width           =   1080
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  实验类型："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1485
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   120
      Picture         =   "frmPatholAntibody_FeedbackUpdate.frx":179A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "    请正确录入抗体反馈信息，以便对其评估。在参考病理号的录入项目中，可根据需要录入参考的病理号。"
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   195
      Width           =   5175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6360
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  反馈意见："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "参考病理号："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   975
      Width           =   1080
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "抗体评价："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3840
      TabIndex        =   9
      Top             =   1485
      Width           =   900
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  反馈医生："
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   2490
      Width           =   1080
   End
End
Attribute VB_Name = "frmPatholAntibody_FeedbackUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mblnIsUpdate As Boolean
Private mblnIsSucceed As Boolean
Private mlngAntibodyId As Long


Public mufgFeedback As ucFlexGrid



Property Get IsSucceed() As Boolean
    IsSucceed = mblnIsSucceed
End Property



Property Get IsUpdate() As Boolean
    IsUpdate = mblnIsUpdate
End Property

Property Let IsUpdate(value As Boolean)
    mblnIsUpdate = value
End Property




Private Sub LoadExperimentType()
    cbxExperimentType.Clear
    
    Call cbxExperimentType.AddItem("0-免疫组化")
    Call cbxExperimentType.AddItem("1-特殊染色")
    Call cbxExperimentType.AddItem("2-分子病理")
    Call cbxExperimentType.AddItem("3-其他")
    
    cbxExperimentType.ListIndex = 0
End Sub


Private Sub LoadAntibodyGrade()
    cbxAntibodyGrade.Clear
    
    Call cbxAntibodyGrade.AddItem("优")
    Call cbxAntibodyGrade.AddItem("良")
    Call cbxAntibodyGrade.AddItem("中")
    Call cbxAntibodyGrade.AddItem("差")
    
    cbxAntibodyGrade.ListIndex = 1
End Sub


Private Function GetTestTypeIndex(ByVal strTestValue As String) As Long
'取得实验类型对应的索引
    GetTestTypeIndex = 0
    
    If strTestValue = "免疫组化" Then
        GetTestTypeIndex = 0
    ElseIf strTestValue = "特殊染色" Then
        GetTestTypeIndex = 1
    ElseIf strTestValue = "分子病理" Then
        GetTestTypeIndex = 2
    Else
        GetTestTypeIndex = 3
    End If
End Function


Public Sub ConfigUpdateFace()
'配置更新录入数据
    With mufgFeedback
        
        txtReferencePatholNum.Text = .Text(.SelectionRow, gstrAntibodyFeedback_参考病理号)
        cbxExperimentType.ListIndex = GetTestTypeIndex(.Text(.SelectionRow, gstrAntibodyFeedback_实验类型))
        cbxAntibodyGrade.Text = .Text(.SelectionRow, gstrAntibodyFeedback_抗体评价)
        dtpFeedbackTime.value = .Text(.SelectionRow, gstrAntibodyFeedback_反馈时间)
        txtFeedbackDoctor.Text = .Text(.SelectionRow, gstrAntibodyFeedback_反馈医生)
        txtFeedbackAdvice.Text = .Text(.SelectionRow, gstrAntibodyFeedback_反馈意见)
    End With
End Sub


Public Function ShowUpdateAntibodyFeedback(ufgFeedback As ucFlexGrid, owner As Form) As Boolean
'显示抗体反馈更新窗口
    
    ShowUpdateAntibodyFeedback = False
    
    Set mufgFeedback = ufgFeedback
    
    Me.Caption = "更新反馈"
    mblnIsUpdate = True
    
    Call CloseProcessHint
    
    Call ConfigUpdateFace
        
    Call Me.Show(1, owner)
    
    ShowUpdateAntibodyFeedback = mblnIsSucceed
End Function


Public Function ShowAddAntibodyFeedback(ByVal lngAntibodyId As Long, ufgFeedback As ucFlexGrid, owner As Form) As Boolean
'显示新增抗体反馈窗口
    ShowAddAntibodyFeedback = False
    
    
    Set mufgFeedback = ufgFeedback
    
    mlngAntibodyId = lngAntibodyId
    
    Me.Caption = "新增反馈"
    mblnIsUpdate = False
    
    dtpFeedbackTime.value = zlDatabase.Currentdate
    txtFeedbackDoctor.Text = UserInfo.姓名
        
    Call CloseProcessHint
    
    Call Me.Show(1, owner)
    
    ShowAddAntibodyFeedback = mblnIsSucceed
End Function


Private Function CheckFeedbackDataIsValid() As String
    CheckFeedbackDataIsValid = ""
    
    If Trim(txtFeedbackAdvice.Text) = "" Then
        CheckFeedbackDataIsValid = "反馈意见不能为空。"
        
        Call txtFeedbackAdvice.SetFocus
        Exit Function
    End If
End Function



Private Function AddFeedbackDataToList(ByVal lngFeedbackId As Long)
'添加抗体记录到显示列表
On Error GoTo errHandle
    AddFeedbackDataToList = ""
    
    Dim lngNewRecordIndex As Long
    
    AddFeedbackDataToList = ""
    
    With mufgFeedback
        lngNewRecordIndex = .NewRow
        
        .Text(lngNewRecordIndex, gstrAntibodyFeedback_ID) = lngFeedbackId
        .Text(lngNewRecordIndex, gstrAntibodyFeedback_参考病理号) = txtReferencePatholNum.Text
        .Text(lngNewRecordIndex, gstrAntibodyFeedback_实验类型) = Trim(Substr(cbxExperimentType.Text, InStr(1, cbxExperimentType.Text, "-") + 1, 20))
        .Text(lngNewRecordIndex, gstrAntibodyFeedback_抗体评价) = cbxAntibodyGrade.Text
        .Text(lngNewRecordIndex, gstrAntibodyFeedback_反馈意见) = txtFeedbackAdvice.Text
        .Text(lngNewRecordIndex, gstrAntibodyFeedback_反馈医生) = txtFeedbackDoctor.Text
        .Text(lngNewRecordIndex, gstrAntibodyFeedback_反馈时间) = dtpFeedbackTime.value
    End With
     
    
Exit Function
errHandle:
    AddFeedbackDataToList = err.Description
End Function




Private Sub cmdFeedback_Cancel_Click()
    mblnIsSucceed = False
    
    Call Me.Hide
End Sub



Private Function UpdateFeedbackInfToList()
'更新抗体列表中的抗体信息
On Error GoTo errHandle
    UpdateFeedbackInfToList = ""
    
    With mufgFeedback
        .Text(.SelectionRow, gstrAntibodyFeedback_参考病理号) = txtReferencePatholNum.Text
        .Text(.SelectionRow, gstrAntibodyFeedback_实验类型) = Trim(Substr(cbxExperimentType.Text, InStr(1, cbxExperimentType.Text, "-") + 1, 20))
        .Text(.SelectionRow, gstrAntibodyFeedback_抗体评价) = cbxAntibodyGrade.Text
        .Text(.SelectionRow, gstrAntibodyFeedback_反馈意见) = txtFeedbackAdvice.Text
        .Text(.SelectionRow, gstrAntibodyFeedback_反馈医生) = txtFeedbackDoctor.Text
        .Text(.SelectionRow, gstrAntibodyFeedback_反馈时间) = dtpFeedbackTime.value
    End With
Exit Function
errHandle:
    UpdateFeedbackInfToList = err.Description
End Function


Private Function UpdateFeedback() As String
'更新抗体反馈信息
On Error GoTo errHandle

    Dim strSql As String
    Dim lngCurFeedbackId As Long
    
    UpdateFeedback = ""
    
    lngCurFeedbackId = mufgFeedback.KeyValue(mufgFeedback.SelectionRow)
    
    strSql = "zl_病理抗体反馈_更新(" & lngCurFeedbackId & ",'" & _
                                txtReferencePatholNum.Text & "'," & _
                                Val(cbxExperimentType.Text) & ",'" & _
                                cbxAntibodyGrade.Text & "'," & _
                                To_Date(dtpFeedbackTime.value) & ",'" & _
                                txtFeedbackDoctor.Text & "','" & _
                                txtFeedbackAdvice.Text & "')"
                                
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
Exit Function
errHandle:
    UpdateFeedback = err.Description
End Function


Private Function NewFeedback(ByVal lngAntibodyId As Long, ByRef lngFeedbackId As Long) As String
'在数据库中新增抗体记录
On Error GoTo errHandle

    Dim strSql As String
    Dim rsReture As ADODB.Recordset
    
    NewFeedback = ""
    
                                
    strSql = "select Zl_病理抗体反馈_新增([1],[2],[3],[4],[5],[6],[7]) as 返回值 from dual"
                                
    Set rsReture = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                lngAntibodyId, _
                                txtReferencePatholNum.Text, _
                                Val(cbxExperimentType.Text), _
                                cbxAntibodyGrade.Text, _
                                CDate(dtpFeedbackTime.value), _
                                txtFeedbackDoctor.Text, _
                                txtFeedbackAdvice.Text)
                                
    If rsReture.RecordCount > 0 Then lngFeedbackId = rsReture!返回值
    
Exit Function
errHandle:
    NewFeedback = err.Description
End Function


Private Sub cmdFeedback_Sure_Click()
On Error GoTo errHandle
    Dim strErr As String
    Dim lngFeedbackId As Long
    
    strErr = CheckFeedbackDataIsValid
    If Trim(strErr) <> "" Then
        Call ShowProcessHint(strErr)
        
        Exit Sub
    End If
    
    If mblnIsUpdate Then
        '更新抗体反馈记录
        strErr = UpdateFeedback()
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
        
        strErr = UpdateFeedbackInfToList()
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
    Else
        '新增抗体反馈记录
        strErr = NewFeedback(mlngAntibodyId, lngFeedbackId)
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
        
        strErr = AddFeedbackDataToList(lngFeedbackId)
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
        
        Call mufgFeedback.LocateRow(mufgFeedback.GridRows - 1)
    End If
    
    mblnIsSucceed = True
    
    Call Me.Hide
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
    mblnIsSucceed = False
    mblnIsUpdate = False
    
    mlngAntibodyId = -1
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadExperimentType
    Call LoadAntibodyGrade
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Exit Sub
End Sub



Private Sub ShowProcessHint(ByVal strHint As String)
'显示处理信息
On Error Resume Next

    txtShow.Text = strHint

    picShow.Visible = True
End Sub


Private Sub CloseProcessHint()
'关闭处理提示
    picShow.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub
