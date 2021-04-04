VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholRequisition_SpeExam 
   Caption         =   "特检申请"
   ClientHeight    =   7980
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10260
   Icon            =   "frmPatholRequisition_SpeExam.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   10260
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtRequestDoctor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   7440
      Width           =   2145
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "添 加(&A)"
      Height          =   400
      Left            =   7320
      TabIndex        =   22
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CheckBox chkPriceState 
      Caption         =   "需补费"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   7560
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退 出(&E)"
      Height          =   400
      Left            =   8640
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame framAntibody 
      Caption         =   "特检项目"
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   9975
      Begin VB.PictureBox picMealClass 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   7080
         ScaleHeight     =   5295
         ScaleWidth      =   2655
         TabIndex        =   7
         Top             =   360
         Width           =   2655
         Begin VB.ListBox lstMealClass 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4050
            Left            =   0
            TabIndex        =   10
            Top             =   720
            Width           =   2655
         End
         Begin VB.ComboBox cbxMealClass 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label labMealClass 
            Caption         =   "套餐类别："
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.TextBox txtAntibodyName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   0
         ToolTipText     =   "根据抗体名称进行快速定位。"
         Top             =   5445
         Width           =   2625
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   4935
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8705
         GridRows        =   21
         IsKeepRows      =   0   'False
         BackColor       =   12648447
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
      Begin VB.Label labFilter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称过滤："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3360
         TabIndex        =   5
         ToolTipText     =   "根据抗体名称进行快速定位。"
         Top             =   5520
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox cbxMaterial 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   3225
      End
      Begin VB.ComboBox cbxSpeExamType 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholRequisition_SpeExam.frx":179A
         Left            =   5520
         List            =   "frmPatholRequisition_SpeExam.frx":179C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   3225
      End
      Begin VB.ComboBox cbxSpeExamDetails 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholRequisition_SpeExam.frx":179E
         Left            =   1080
         List            =   "frmPatholRequisition_SpeExam.frx":17A0
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   3225
      End
      Begin VB.Label labMaterial 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "材块编号："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   900
      End
      Begin VB.Label labSpeExamType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特检类型："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4560
         TabIndex        =   20
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Width           =   200
      End
      Begin VB.Label Label5 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8760
         TabIndex        =   18
         Top             =   240
         Width           =   200
      End
      Begin VB.Label labDescription 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申请描述："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4560
         TabIndex        =   17
         Top             =   810
         Width           =   900
      End
      Begin VB.Label labSpeexamDetails 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特检细目："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   900
      End
   End
   Begin MSComCtl2.DTPicker dtpRequestTime 
      Height          =   300
      Left            =   720
      TabIndex        =   24
      Top             =   7440
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   74317827
      CurrentDate     =   40646.4399652778
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   8640
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labRequestDoctor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医师："
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3000
      TabIndex        =   26
      Top             =   7500
      Width           =   540
   End
   Begin VB.Label labRequestTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "时间："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   7500
      Width           =   540
   End
End
Attribute VB_Name = "frmPatholRequisition_SpeExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mufgParentRequest As ucFlexGrid
Private mufgParentContext As ucFlexGrid

Private mblnIsBuZuo As Boolean
Private mfrmOwner As Form

Private mrsMealLink As New ADODB.Recordset

Private mlngCurRequestId As Long
Private mlngPatholAdviceId As Long


Public blnIsOk As Boolean


Public Function ShowSpeExamRequestWindow(ufgParentRequestGrid As ucFlexGrid, ufgParentContextGrid As ucFlexGrid, _
    ByVal lngPatholAdviceId As Long, ByVal lngRequestId As Long, owner As Form, Optional ByVal blnIsBuZuo As Boolean = False) As Boolean
'显示特检申请窗口
    Set mufgParentRequest = ufgParentRequestGrid
    Set mufgParentContext = ufgParentContextGrid

    Set mfrmOwner = owner
    
    mlngPatholAdviceId = lngPatholAdviceId
    mlngCurRequestId = lngRequestId
    mblnIsBuZuo = blnIsBuZuo
    blnIsOk = False
        
    '载入材块信息
    Call LoadMaterialInf
    
    dtpRequestTime.value = zlDatabase.Currentdate
    txtRequestDoctor.Text = UserInfo.姓名
    
    If lngRequestId > 0 Then
        Select Case ufgParentRequestGrid.Text(ufgParentRequestGrid.SelectionRow, gstrRequisition_申请类型)
            Case "免疫组化"
                cbxSpeExamType.ListIndex = 0
            Case "特殊染色"
                cbxSpeExamType.ListIndex = 1
            Case "分子病理"
                cbxSpeExamType.ListIndex = 2
        End Select
        
        Call LoadSpeExamDetails(Val(cbxSpeExamType.Text))
        
        Select Case ufgParentRequestGrid.Text(ufgParentRequestGrid.SelectionRow, gstrRequisition_申请细目)
            Case "鉴别", "荧光"
                cbxSpeexamDetails.ListIndex = 0
            Case "多药耐药", "普通"
                cbxSpeexamDetails.ListIndex = 1
        End Select
        
    End If
    
    '调整特检申请界面
    Call AdjustSpeExamFace(lngRequestId <= 0)
    
    Call Me.Show(1, owner)
End Function


Private Sub AdjustSpeExamFace(ByVal blnIsNewRequest As Boolean)
'调整特检界面
    labRequestTime.Enabled = blnIsNewRequest
    dtpRequestTime.Enabled = blnIsNewRequest
    
    labSpeExamType.Enabled = blnIsNewRequest
    cbxSpeExamType.Enabled = blnIsNewRequest
    
    labSpeexamDetails.Enabled = blnIsNewRequest
    cbxSpeexamDetails.Enabled = blnIsNewRequest
    
    labRequestDoctor.Enabled = blnIsNewRequest
    txtRequestDoctor.Enabled = blnIsNewRequest
    
    labDescription.Enabled = blnIsNewRequest
    txtDescription.Enabled = blnIsNewRequest
    
    cbxSpeExamType.BackColor = IIf(blnIsNewRequest, vbWhite, Me.BackColor)
    cbxSpeexamDetails.BackColor = IIf(blnIsNewRequest, vbWhite, Me.BackColor)
    txtDescription.BackColor = IIf(blnIsNewRequest, vbWhite, Me.BackColor)
End Sub


Private Sub AdjustFace()
    framAntibody.Height = Me.Height - framAntibody.Top - cmdExit.Height - 800
    framAntibody.Width = Me.Width - (framAntibody.Left * 2) - 120
    
    ufgData.Left = 120
    ufgData.Top = 240
    
    ufgData.Height = framAntibody.Height - txtAntibodyName.Height - 480
    ufgData.Width = framAntibody.Width - picMealClass.Width - 360
    
    picMealClass.Left = ufgData.Left + ufgData.Width + 120
    picMealClass.Top = ufgData.Top
    picMealClass.Height = ufgData.Height + txtAntibodyName.Height + 120
    
    lstMealClass.Height = picMealClass.Height - lstMealClass.Top - txtAntibodyName.Height
        
    
    labFilter.Left = 120
    
    
    txtAntibodyName.Left = labFilter.Left + labFilter.Width + 60
    txtAntibodyName.Top = ufgData.Top + ufgData.Height + 120
    txtAntibodyName.Width = ufgData.Width - txtAntibodyName.Left + 120
    
    labFilter.Top = txtAntibodyName.Top + 60
    
    
    cmdExit.Left = framAntibody.Width - cmdExit.Width + framAntibody.Left
    cmdExit.Top = framAntibody.Top + framAntibody.Height + 120
    
    cmdApply.Left = cmdExit.Left - cmdApply.Width - 120
    cmdApply.Top = cmdExit.Top
    
    
    chkPriceState.Left = cmdApply.Left - chkPriceState.Width - 120
    chkPriceState.Top = cmdExit.Top + 90
    
    labRequestTime.Left = 120
    labRequestTime.Top = chkPriceState.Top
    
    dtpRequestTime.Left = labRequestTime.Left + labRequestTime.Width + 60
    dtpRequestTime.Top = cmdExit.Top + 30
    
    labRequestDoctor.Left = dtpRequestTime.Left + dtpRequestTime.Width + 240
    labRequestDoctor.Top = labRequestTime.Top
    
    txtRequestDoctor.Left = labRequestDoctor.Left + labRequestDoctor.Width + 60
    txtRequestDoctor.Top = cmdExit.Top + 30
End Sub



Private Sub LoadMaterialInf()
'载入材块信息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    
    strSQL = "select 材块ID,序号,标本名称 from 病理取材信息 where 病理医嘱ID=[1] and 确认状态=1"
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatholAdviceId)
    
    Call cbxMaterial.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsData.EOF
        Call cbxMaterial.AddItem(Nvl(rsData!材块id) & " (材块号：" & rsData!序号 & " 标本名：" & rsData!标本名称 & ")")

        Call rsData.MoveNext
    Loop
    
End Sub



Private Sub LoadSpeExamType()
'载入抗体类型
    cbxSpeExamType.Clear
    
    Call cbxSpeExamType.AddItem("0-免疫组化")
    Call cbxSpeExamType.AddItem("1-特殊染色")
    Call cbxSpeExamType.AddItem("2-分子病理")
    
    cbxSpeExamType.ListIndex = 0
End Sub



Private Sub LoadSpeExamDetails(ByVal lngSpeExamType As Long)
'载入特检明细
    cbxSpeexamDetails.Clear
    
'    Call cbxSpeExamDetails.AddItem("")
    
    If lngSpeExamType = TSpeexamType.stMianyi Then
        Call cbxSpeexamDetails.AddItem("1-免疫(鉴别)")
        Call cbxSpeexamDetails.AddItem("2-免疫(多药耐药)")
        
        cbxSpeexamDetails.ListIndex = 1
    ElseIf lngSpeExamType = TSpeexamType.stFenzi Then
        Call cbxSpeexamDetails.AddItem("1-分子(荧光)")  '对应 3
        Call cbxSpeexamDetails.AddItem("2-分子(普通)")  '对应 4
        
        cbxSpeexamDetails.ListIndex = 0
    End If
End Sub



Private Sub LoadMealLinkData()
'载入套餐关联数据
    Dim strSQL As String
    
    '读取套餐关联数据
    strSQL = "select a.套餐ID, a.套餐名称, b.抗体id, b.抗体顺序 from 病理套餐信息 a, 病理套餐关联 b where a.套餐id=b.套餐id"
    Set mrsMealLink = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
End Sub


Private Sub LoadAntibodyMeal(ByVal strMealClass As String)
'载入抗体套餐
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    '读取套餐数据
    strSQL = "select 套餐名称 from 病理套餐信息 " & IIf(strMealClass <> "", " where 套餐类别=[1]", "") & " order by 套餐名称"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMealClass)
    
    Call lstMealClass.Clear
    
    Call lstMealClass.AddItem("")
    If rsData.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsData.EOF
        Call lstMealClass.AddItem(Nvl(rsData!套餐名称))
        rsData.MoveNext
    Loop
    
End Sub


Private Sub InitAntibodyList()
'初始化抗体信息列表
    Dim strTemp As String
    

    
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("特检抗体列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrRequestAntibodyCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    '禁止右键弹出列表配置窗口
    ufgData.IsEjectConfig = False
        '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.DefaultColNames = gstrRequestAntibodyCols
    ufgData.ColConvertFormat = gstrRequestAntibodyConvertFormat
End Sub



Private Sub QueryAntibodyData()
'查询抗体数据
    Dim strSQL As String
    
    strSQL = "select a.抗体id, a.抗体名称,a.使用人份,a.已用人份,a.生产日期,a.有效期,a.过期日期, '' as 项目顺序 " & _
                " from 病理抗体信息 a where a.使用状态=1 order by a.抗体名称 "
                
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call ufgData.RefreshData
End Sub


Private Sub GetMealAntibodyIds(ByVal strMealName As String, ByRef strAntibodyIds As String, ByRef strAntibodyOrder As String)
'取得套餐所属的抗体Id串
    
    strAntibodyIds = ""
    strAntibodyOrder = ""
    
    
    mrsMealLink.Filter = "套餐名称='" & strMealName & "'"
    If mrsMealLink.RecordCount <= 0 Then Exit Sub
    
    Do While Not mrsMealLink.EOF
        If strAntibodyIds <> "" Then strAntibodyIds = strAntibodyIds & " or "
        strAntibodyIds = strAntibodyIds & "抗体Id=" & mrsMealLink!抗体ID
        strAntibodyOrder = strAntibodyOrder & mrsMealLink!抗体ID & ":" & _
            String(5 - Len("" & mrsMealLink!套餐ID & ""), "0") & mrsMealLink!套餐ID & String(5 - Len("" & mrsMealLink!抗体顺序 & ""), "0") & mrsMealLink!抗体顺序 & ";"
        
        Call mrsMealLink.MoveNext
    Loop

End Sub


Private Sub LoadAntibodyDataToFace(ByVal strAntibodyMeal As String)
'读取抗体信息到界面显示
    Dim strCurMealAntibodyIds As String
    Dim strCurAntibodyOrders As String
    Dim strCurAntibodyOrder As String
    Dim i As Long
    
    '取得当前套餐所属的套餐ID
    If strAntibodyMeal <> "" Then
        Call GetMealAntibodyIds(strAntibodyMeal, strCurMealAntibodyIds, strCurAntibodyOrders)
        ufgData.AdoFilter = IIf(strCurMealAntibodyIds <> "", strCurMealAntibodyIds, "抗体ID=-1")
    Else
        ufgData.AdoFilter = ""
    End If
    
    Call ufgData.RefreshData
    
    '写入当前套餐的抗体顺序
    If strCurAntibodyOrders = "" Then Exit Sub
    
    On Error Resume Next
    For i = 1 To ufgData.GridRows - 1
        strCurAntibodyOrder = ufgData.KeyValue(i)
        strCurAntibodyOrder = Mid(strCurAntibodyOrders, InStr(strCurAntibodyOrders, strCurAntibodyOrder & ":") + Len(strCurAntibodyOrder) + 1, 100)
        strCurAntibodyOrder = Mid(strCurAntibodyOrder, 1, InStr(strCurAntibodyOrder, ";") - 1)
        
        '设置套餐下的抗体顺序
        ufgData.Text(i, gstrRequestAntibody_项目顺序) = strCurAntibodyOrder
    Next i
    
    '按套餐的抗体顺序排序
    Call ufgData.Sort(ufgData.GetColIndex(gstrRequestAntibody_项目顺序))
End Sub



Private Function GetSpecimenAntibodyIds(ByVal strMaterialId As String)
'获取标本对应的抗体Id
    Dim i As Long
    Dim strIds As String
    
    'strIds 内容形式为",asf,aat,aft,bbe,"
    
    strIds = ""
    For i = 1 To mufgParentContext.GridRows - 1
        If Not mufgParentContext.IsEmptyKey(i) Then
            If Val(mufgParentContext.Text(i, gstrRequest_SpeExam_材块号)) = Val(strMaterialId) Then
                strIds = strIds & "," & mufgParentContext.Text(i, gstrRequest_SpeExam_抗体名称)
            End If
        End If
    Next i
    
    If strIds <> "" Then strIds = strIds & ","
    
    GetSpecimenAntibodyIds = strIds
End Function


Private Function HideSpecimenAntibody(ByVal strSpecimenAntibodyIds As String)
'隐藏标本已经选择的抗体
    Dim i As Long
    
    For i = 1 To ufgData.GridRows - 1
        If UCase(strSpecimenAntibodyIds) Like "*," & UCase(ufgData.Text(i, gstrRequestAntibody_抗体名称)) & ",*" Then
            ufgData.RowHidden(i) = True
        End If
    Next i
End Function


Private Sub cbxMaterial_Click()
On Error GoTo ErrHandle
    Dim strSpecimenAntibodyIds As String
    
    '恢复列表所隐藏的数据
    Call ufgData.RestoreList


    
'    If mblnIsBuZuo Then
        strSpecimenAntibodyIds = GetSpecimenAntibodyIds(GetSelectMaterialNum)
    
        '隐藏对应的抗体
        Call HideSpecimenAntibody(strSpecimenAntibodyIds)
'    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbxMealClass_Click()
On Error GoTo ErrHandle

    Call LoadAntibodyMeal(cbxMealClass.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxSpeExamType_Click()
On Error GoTo ErrHandle
    '载入特检细目
    Call LoadSpeExamDetails(Val(cbxSpeExamType.Text))
    
    Exit Sub
ErrHandle:
    If ErrCenter() = True Then Resume
End Sub

Private Sub cmdApply_Click()
'保存特检申请
On Error GoTo ErrHandle
    '如果数据无效，则直接退出，不需要再进行提示
    If Not CheckDataIsValid Then Exit Sub
    
    '判断申请明细列表是否为特检项目明细列表
    If mufgParentContext.GetColIndex(gstrRequest_SpeExam_抗体名称) < 0 Then
        mufgParentContext.ColNames = gstrRequest_SpeExam_Cols
        mufgParentContext.ColConvertFormat = gstrRequest_SpeExamConvertFormat
        
        
        '切换主界面的控制界面
        Call mfrmOwner.ChangeControlFace(0)
    End If
    
    '保存特检申请
    Call SaveSpeExamRequest
    
    blnIsOk = True
    
    If MsgBoxD(Me, "操作已完成，是否继续添加？", vbYesNo, Me.Caption) = vbNo Then
        Call Me.Hide
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
'退出窗口
On Error Resume Next
    blnIsOk = False
    Call Me.Hide
End Sub



Private Function CheckDataIsValid() As Boolean
'检查录入数据是否有效
    CheckDataIsValid = False
    
    '判断是否选择了材块
    If Trim(cbxMaterial.Text) = "" Then
        Call MsgBoxD(Me, "请选择需要进行特检处理的材块。", vbInformation, Me.Caption)
        cbxMaterial.SetFocus
        
        Exit Function
    End If
    
    
    If mlngCurRequestId <= 0 Then
        '判断是否选择了特检类型  (只有新申请才需要进行判断)
        If Trim(cbxSpeExamType.Text) = "" Then
            Call MsgBoxD(Me, "请选择特检的处理类型。", vbInformation, Me.Caption)
            cbxSpeExamType.SetFocus
            
            Exit Function
        End If
    End If
    
    '判断是否选择抗体数据
    If Not ufgData.IsCheckedRow Then
        Call MsgBoxD(Me, "请选择需要使用的抗体数据。", vbInformation, Me.Caption)
        ufgData.SetFocus
        
        Exit Function
    End If
    
    CheckDataIsValid = True

End Function



Private Sub SaveSpeExamRequest()
'保存特检申请
    Dim lngNewRow As Long
    
    Dim i As Long
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnContinueAdd As Boolean
    Dim lngSpeexamDetails As Long
    Dim blnIsInvalidDate As Boolean
    Dim blnIsInvalidCount As Boolean
    
    lngSpeexamDetails = 0
    
    '获取当前特检细目
    Select Case Val(cbxSpeExamType.Text)
        Case 0
            lngSpeexamDetails = Val(cbxSpeexamDetails.Text)
        Case 1
            lngSpeexamDetails = 0
        Case 2
            lngSpeexamDetails = IIf(Val(cbxSpeexamDetails.Text) > 0, Val(cbxSpeexamDetails.Text) + 2, 0)
    End Select
        
        
        
    If mlngCurRequestId <= 0 Then
        
        '添加检查申请信息
        strSQL = "select Zl_病理申请_新增([1],[2],[3],[4],[5],[6],[7]) as 返回值 from dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                mlngPatholAdviceId, _
                                                txtRequestDoctor.Text, _
                                                CDate(dtpRequestTime.value), _
                                                Val(cbxSpeExamType.Text), _
                                                IIf(chkPriceState.value <> 0, 1, 0), _
                                                lngSpeexamDetails, _
                                                txtDescription.Text)
                                                
        If rsData.RecordCount <= 0 Then
            Call err.Raise(0, "SaveSpeExamRequest", "未成功获取新增后的申请ID,处理失败。")
            Exit Sub
        End If

        '设置界面信息
        lngNewRow = mufgParentRequest.NewRow
        
        mufgParentRequest.Text(lngNewRow, gstrRequisition_申请ID) = rsData!返回值
        mufgParentRequest.Text(lngNewRow, gstrRequisition_申请人) = txtRequestDoctor.Text
        mufgParentRequest.Text(lngNewRow, gstrRequisition_申请类型) = Trim(Substr(cbxSpeExamType.Text, InStr(1, cbxSpeExamType.Text, "-") + 1, 10))
        mufgParentRequest.Text(lngNewRow, gstrRequisition_申请细目) = Decode(lngSpeexamDetails, 1, "鉴别", 2, "多药耐药", 3, "荧光", 4, "普通", "无")
        mufgParentRequest.Text(lngNewRow, gstrRequisition_补费状态) = IIf(chkPriceState.value <> 0, "需补费", "无")
        mufgParentRequest.Text(lngNewRow, gstrRequisition_申请时间) = dtpRequestTime.value
        mufgParentRequest.Text(lngNewRow, gstrRequisition_申请时间) = dtpRequestTime.value
        mufgParentRequest.Text(lngNewRow, gstrRequisition_申请描述) = txtDescription.Text
        mufgParentRequest.Text(lngNewRow, gstrRequisition_当前状态) = "已申请"
                                                
        mlngCurRequestId = Val(Nvl(rsData!返回值))
        
        '定位到新增行
        Call mufgParentRequest.LocateRow(lngNewRow)
        
        '清除原有申请项目数据
        Call mufgParentContext.ClearListData
    End If
    
    
    
    '添加特检申请项目
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetRowCheck(i) And Not ufgData.RowHidden(i) Then
        
            blnContinueAdd = True
            
            blnIsInvalidCount = Val(ufgData.Text(i, gstrRequestAntibody_使用人份)) <= Val(ufgData.Text(i, gstrRequestAntibody_已用人份))
            blnIsInvalidDate = zlDatabase.Currentdate > CDate(IIf(ufgData.Text(i, gstrRequestAntibody_过期日期) = "", "3000-01-01", ufgData.Text(i, gstrRequestAntibody_过期日期))) _
                    Or zlDatabase.Currentdate > DateAdd("m", _
                        Val(IIf(ufgData.Text(i, gstrRequestAntibody_有效期) = "", 2400, ufgData.Text(i, gstrRequestAntibody_有效期))), _
                        CDate(ufgData.Text(i, gstrRequestAntibody_生产日期)))
            
            
            '判断是否存在使用人份
            If blnIsInvalidCount Or blnIsInvalidDate Then
                If MsgBoxD(Me, "抗体 [" & ufgData.Text(i, gstrRequestAntibody_抗体名称) & "]" & _
                                IIf(blnIsInvalidCount, "已无可用人份，", "") & _
                                IIf(blnIsInvalidDate, "已过有效期，", "") & "是否继续添加该项目？", vbYesNo, Me.Caption) <> vbYes Then
                    blnContinueAdd = False

                End If
            End If
                        
                        

            
            If blnContinueAdd Then
                strSQL = "select Zl_病理申请_特检项目_新增([1],[2],[3],[4],[5],[6],[7],[8],[9]) as 返回值 from dual"
                Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                        mlngPatholAdviceId, _
                                                        GetSelectedMaterialId, _
                                                        mlngCurRequestId, _
                                                        Val(ufgData.KeyValue(i)), _
                                                        Val(cbxSpeExamType.Text), _
                                                        lngSpeexamDetails, _
                                                        IIf(chkPriceState.value <> 0, 1, 0), _
                                                        ufgData.Text(i, gstrRequestAntibody_项目顺序), _
                                                        IIf(mblnIsBuZuo, 1, 0))
                                                        
                If rsData.RecordCount <= 0 Then
                    Call err.Raise(0, "SaveSpeExamRequest", "未成功获取新增后的特检项目ID,处理失败。")
                    Exit Sub
                End If
                
                '设置界面信息
                lngNewRow = mufgParentContext.NewRow
                
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_ID) = rsData!返回值
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_标本名称) = GetSelectSpecimenName
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_材块号) = GetSelectMaterialNum
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_抗体名称) = ufgData.Text(i, gstrRequestAntibody_抗体名称)
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_制作类型) = IIf(mblnIsBuZuo, "补做", "常规")
                mufgParentContext.Text(lngNewRow, gstrRequest_SpeExam_当前状态) = "已申请"
                
                Call mufgParentContext.LocateRow(lngNewRow)
                
                ufgData.RowHidden(i) = True
            End If

        End If
    Next i
    
End Sub


Private Function GetSelectedMaterialId()
'取得当前选择的材块ID
    GetSelectedMaterialId = Substr(cbxMaterial.Text, 1, InStr(1, cbxMaterial.Text, "(") - 1)
End Function


Private Function GetSelectSpecimenName()
'取得当前选择的标本名称
    Dim strMaterialInf As String
    Dim strReplace As String
    
    GetSelectSpecimenName = ""
    If Trim(cbxMaterial.Text) = "" Then Exit Function
    
    strMaterialInf = cbxMaterial.Text
    
    strReplace = Left(strMaterialInf, InStr(1, strMaterialInf, "标本名：") + 3)
    
    strMaterialInf = Replace(strMaterialInf, strReplace, "")
    
    
    GetSelectSpecimenName = Mid(strMaterialInf, 1, Len(strMaterialInf) - 1)
End Function



Private Function GetSelectMaterialNum()
'取得当前选择的材块号
    Dim strMaterialInf As String
    Dim strReplace As String
    
    GetSelectMaterialNum = ""
    If Trim(cbxMaterial.Text) = "" Then Exit Function
    
    strMaterialInf = cbxMaterial.Text
    
    strReplace = Mid(strMaterialInf, 1, InStr(1, strMaterialInf, "(材块号：") + 4)
    
    strMaterialInf = Replace(strMaterialInf, strReplace, "")
    
    
    GetSelectMaterialNum = Mid(strMaterialInf, 1, InStr(strMaterialInf, " 标本名：") - 1)
End Function

Private Sub cmdSure_Click()
'保存特检申请
On Error GoTo ErrHandle
    Call cmdApply_Click
    
    If blnIsOk Then Call Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
'初始化变量
    mlngCurRequestId = -1
    mlngPatholAdviceId = -1
    mblnIsBuZuo = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    '初始化抗体显示列表
    Call InitAntibodyList
    
    '载入特检类型
    Call LoadSpeExamType
    
    '载入特检细目
    Call LoadSpeExamDetails(Val(cbxSpeExamType.Text))
    
    '查询抗体信息
    Call QueryAntibodyData
    
    '载入套餐关联数据
    Call LoadMealLinkData
    
    '载入套餐分类
    Call LoadMealClass
    
    '载入套餐信息
    Call LoadAntibodyMeal(cbxMealClass.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Exit Sub
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    '关闭窗口时保存列表配置
     zlDatabase.SetPara "特检抗体列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
    
    Set mrsMealLink = Nothing
End Sub


Private Sub LoadMealClass()
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select distinct 套餐类别 from 病理套餐信息"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cbxMealClass.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    cbxMealClass.AddItem ""
    While Not rsData.EOF
        If Nvl(rsData!套餐类别) <> "" Then cbxMealClass.AddItem Nvl(rsData!套餐类别)
        rsData.MoveNext
    Wend
End Sub

Private Sub lstMealClass_Click()
On Error GoTo ErrHandle
   
    Call LoadAntibodyDataToFace(lstMealClass.Text)
    
    Call cbxMaterial_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtAntibodyName_Change()
''根据抗体名称过滤
'On Error GoTo errHandle
'    Dim lngFindIndex As Long
'    Dim i As Long
'    Dim strTxt() As String
'
'    If txtAntibodyName.Text = "" Then Exit Sub
'
'    strTxt() = Split(txtAntibodyName.Text, " ")
'
'    For i = LBound(strTxt) To UBound(strTxt)
'        lngFindIndex = ufgData.FindRowIndex(strTxt(i), gstrRequestAntibody_抗体名称, True)
'        If lngFindIndex > 0 Then
'            Call ufgData.SetRowChecked(lngFindIndex, True, csSystem)
'            Call ufgData.LocateRow(lngFindIndex)
'        End If
'    Next i
'
'
''    If Trim(txtAntibodyName.Text) = "" Then Exit Sub
'
''    lngFindIndex = ufgData.FindRowIndex(txtAntibodyName.Text, gstrRequestAntibody_抗体名称)
''
''    If lngFindIndex > 0 Then Call ufgData.LocateRow(lngFindIndex)
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ShowAntibodyInf(ByVal lngAntibodyRow As Long)
'显示抗体明细信息
    Dim frmAntibodyInf As New frmPatholRequisition_AntibodyInf
    On Error GoTo errFree
        Call frmAntibodyInf.ShowAntibodyInf(ufgData.KeyValue(lngAntibodyRow), Me)
errFree:
    Call Unload(frmAntibodyInf)
    Set frmAntibodyInf = Nothing
    
End Sub



Private Sub txtAntibodyName_KeyPress(KeyAscii As Integer)
'根据抗体名称过滤
On Error GoTo ErrHandle
    Dim lngFindIndex As Long
    Dim i As Long
    Dim strTxt() As String
    
    If txtAntibodyName.Text = "" Then Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    strTxt() = Split(txtAntibodyName.Text, " ")
    
    For i = LBound(strTxt) To UBound(strTxt)
        lngFindIndex = ufgData.FindRowIndex(strTxt(i), gstrRequestAntibody_抗体名称, True)
        If lngFindIndex > 0 Then
            Call ufgData.SetRowCheck(lngFindIndex, True)
            Call ufgData.LocateRow(lngFindIndex)
        End If
    Next i
    
    
'    If Trim(txtAntibodyName.Text) = "" Then Exit Sub
    
'    lngFindIndex = ufgData.FindRowIndex(txtAntibodyName.Text, gstrRequestAntibody_抗体名称)
'
'    If lngFindIndex > 0 Then Call ufgData.LocateRow(lngFindIndex)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Call ShowAntibodyInf(Row)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnDblClick()
'双击显示抗体明细
On Error GoTo ErrHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then Exit Sub
    
    Call ShowAntibodyInf(ufgData.SelectionRow)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

