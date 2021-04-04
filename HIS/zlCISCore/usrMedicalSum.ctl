VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.UserControl usrMedicalSum 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   LockControls    =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   7935
   Begin VB.Frame fra 
      BackColor       =   &H80000005&
      Height          =   2280
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7110
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   0
         Left            =   60
         ScaleHeight     =   330
         ScaleWidth      =   5550
         TabIndex        =   1
         Top             =   150
         Width           =   5550
         Begin MSComctlLib.Toolbar cbr 
            Height          =   345
            Index           =   0
            Left            =   465
            TabIndex        =   2
            Top             =   0
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   609
            ButtonWidth     =   1349
            ButtonHeight    =   609
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ils16"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "收集"
                  Key             =   "收集"
                  Object.ToolTipText     =   "收集各个体检项目的小结"
                  Object.Tag             =   "收集"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "评估"
                  Key             =   "评估"
                  Object.ToolTipText     =   "按评估规则评估结论"
                  Object.Tag             =   "评估"
                  ImageKey        =   "new"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "清空"
                  Key             =   "清空"
                  Object.ToolTipText     =   "清空下面的所有结论"
                  Object.Tag             =   "清空"
                  ImageKey        =   "cls"
               EndProperty
            EndProperty
         End
         Begin VB.Label picTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结论"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   3
            Top             =   75
            Width           =   360
         End
      End
      Begin zl9CISCore.VsfGrid vsf 
         Height          =   1695
         Left            =   165
         TabIndex        =   4
         Top             =   525
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2990
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7260
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":0000
            Key             =   "cls"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":0296
            Key             =   "search"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":6AF8
            Key             =   "new"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":D35A
            Key             =   "newadvice"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":13BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":1A41E
            Key             =   "SelAll"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":20C80
            Key             =   "SelDel"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BackColor       =   &H80000005&
      Height          =   1290
      Index           =   1
      Left            =   30
      TabIndex        =   9
      Top             =   2325
      Width           =   6960
      Begin VB.TextBox rtb 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Top             =   480
         Width           =   1170
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   1
         Left            =   15
         ScaleHeight     =   330
         ScaleWidth      =   4695
         TabIndex        =   10
         Top             =   105
         Width           =   4695
         Begin MSComctlLib.Toolbar cbr 
            Height          =   345
            Index           =   1
            Left            =   570
            TabIndex        =   11
            Top             =   15
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   609
            ButtonWidth     =   1349
            ButtonHeight    =   609
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ils16"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "收集"
                  Key             =   "收集"
                  Object.ToolTipText     =   "收集各体检项目的建议内容"
                  Object.Tag             =   "收集"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "生成"
                  Key             =   "生成"
                  Object.ToolTipText     =   "按上面的结论生成缺省建议"
                  Object.Tag             =   "生成"
                  ImageKey        =   "newadvice"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "清空"
                  Key             =   "清空"
                  Object.ToolTipText     =   "清空下面的建议内容"
                  Object.Tag             =   "清空"
                  ImageKey        =   "cls"
               EndProperty
            EndProperty
         End
         Begin VB.Label picTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "建议"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   12
            Top             =   75
            Width           =   360
         End
      End
   End
   Begin VB.Frame fraOther 
      BackColor       =   &H80000005&
      Height          =   615
      Left            =   165
      TabIndex        =   5
      Top             =   4320
      Width           =   7110
      Begin VB.CommandButton cmd 
         Caption         =   "复查项目"
         Height          =   350
         Left            =   2265
         TabIndex        =   16
         Top             =   165
         Visible         =   0   'False
         Width           =   1100
      End
      Begin MSMask.MaskEdBox msk 
         Height          =   240
         Left            =   1245
         TabIndex        =   14
         Top             =   225
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         ForeColor       =   255
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "YYYY-MM-DD"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   4605
         TabIndex        =   13
         Top             =   210
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "复查时间:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   1125
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "随访期限:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3465
         TabIndex        =   6
         Top             =   225
         Width           =   1140
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4590
         X2              =   5160
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1260
         X2              =   2175
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "月"
         Height          =   180
         Index           =   1
         Left            =   5220
         TabIndex        =   8
         Top             =   255
         Width           =   180
      End
   End
End
Attribute VB_Name = "usrMedicalSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstr挂号单 As String                    '外界传入
Private mlng病历id As Long                      '外界传入
Private mlng医嘱id As Long                      '外界传入
Private mlng病人id As Long                      '外界传入

Private mblnMode As Boolean '为真是表示是用户进行的编辑，这时才赋值
Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String
Private mobjParentObject As Object
Private mrsCallBack As New ADODB.Recordset

Private Enum mCol
    结论描述 = 1
    异常结果
    疾病
    诊断建议
End Enum


Public Function ShowFilterDiagBox(ByVal frmParent As Object, _
                                    ByVal objCmd As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True, _
                                    Optional ByVal blnMuli As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能;显示文本输入选择列表(只用于表格控件)
    '------------------------------------------------------------------------------------------------------------------

    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo errHand


    If rsData.BOF Then
        If blnPrompt Then MsgBox "没有找到相匹配的结果！", , gstrSysName
        Exit Function                            '没有结果，直接返回
    End If
    If rsData.RecordCount = 1 And blnFilter Then GoTo Over                    '因为是输入查找，如果只有一条，则直接返回
        
    Call ClientToScreen(objCmd.hWnd, objPoint)
    lngX = objPoint.x * Screen.TwipsPerPixelX
    lngY = objPoint.y * Screen.TwipsPerPixelY + objCmd.Height

    If frmSelectDialog.ShowSelect(Nothing, 2, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, , , strSavePath, , False, blnMuli) Then GoTo Over
    
    Exit Function
    
Over:
    
    Set rsResult = rsData
    
    ShowFilterDiagBox = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitControl() As Boolean
    
    With vsf
    
        .Cols = 0
        .NewColumn "", 255
        .NewColumn "结论描述", 2400, 1, "...", 1
        .NewColumn "异常结果", 3000, 1, , 1
        .NewColumn "疾病", 600, 1, , 1
        .NewColumn "诊断建议", 15, 1, , 1
        .FixedCols = 1
        
        .ColDataType(mCol.疾病) = flexDTBoolean
        
        .TextMatrix(1, mCol.结论描述) = "未见异常"
        
        .Body.Appearance = flexXPThemes
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
                
    End With
    
    Set mrsCallBack = Nothing
End Function

Public Property Set ParentObject(vData As Object)
    Set mobjParentObject = vData
End Property

Public Property Get ParentObject() As Object
    Set ParentObject = mobjParentObject
End Property

Private Property Let Modified(vData As Boolean)
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    mobjParentObject.Modified = vData
    
End Property

Private Property Get Modified() As Boolean
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    Modified = mobjParentObject.Modified
    
End Property



'公共方法、属性
Public Sub SetgcnOracle()
    '------------------------------------------------------------------------------------------------------------------
    '接口过程
    '------------------------------------------------------------------------------------------------------------------
    Call InitCommon(gcnOracle)
    
End Sub

Public Property Get DispMode() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '接口过程:是否为显示模式
    '------------------------------------------------------------------------------------------------------------------
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    mDispMode = New_DispMode
    
    ShowUsrControl mlng医嘱id, Not mDispMode
    PropertyChanged "DispMode"
    
    If mDispMode Then
        vsf.Body.Editable = flexEDNone
        
        rtb.Locked = True
        
        cbr(0).Buttons("收集").Enabled = False
        cbr(0).Buttons("评估").Enabled = False
        cbr(0).Buttons("清空").Enabled = False
                        
        cbr(1).Buttons("收集").Enabled = False
        cbr(1).Buttons("生成").Enabled = False
        cbr(1).Buttons("清空").Enabled = False
        
        cbr(0).Visible = False
        cbr(1).Visible = False
        
        fraOther.Enabled = False
        
    Else
        cbr(0).Visible = True
        cbr(1).Visible = True
        
        fraOther.Enabled = True
    End If
    
End Property
Public Property Let 挂号单(ByVal New_挂号单 As String)
    '------------------------------------------------------------------------------------------------------------------
    '设置挂号单
    '------------------------------------------------------------------------------------------------------------------
    
    mstr挂号单 = New_挂号单
    
End Property

Public Property Let 病人id(ByVal New_病人id As String)
    '------------------------------------------------------------------------------------------------------------------
    '设置挂号单
    '------------------------------------------------------------------------------------------------------------------
    
    mlng病人id = New_病人id
    
End Property

Public Property Get ID病人病历() As Long
    '------------------------------------------------------------------------------------------------------------------
    '返回病人病历ID
    '------------------------------------------------------------------------------------------------------------------
    
    ID病人病历 = mlng病历id
    
End Property

Public Property Let ID病人病历(ByVal New_ID病人病历 As Long)
    '------------------------------------------------------------------------------------------------------------------
    '设置病人病历ID,并检查该病历是不是存在
    '------------------------------------------------------------------------------------------------------------------
    
    mlng病历id = New_ID病人病历
    ShowUsrControl mlng医嘱id, Not mDispMode
    
End Property

Public Sub SetDiagItem(ByVal New_医嘱ID As Long, ByVal New_发送号)
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    mlng医嘱id = New_医嘱ID
    
End Sub

Public Property Get Get医嘱id() As Long
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Get医嘱id = mlng医嘱id
        
End Property

Public Property Get Text() As String
    '------------------------------------------------------------------------------------------------------------------
    '为每一个控件加上文本转储属性
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngLoop As Long
    Dim strTmp As String
    Dim intCount As Integer
    
    On Error GoTo errHand
    
    '转储结论记录
    intCount = 0
    strTmp = strTmp & "一、结论：" & vbCrLf
    For lngLoop = 1 To vsf.Rows - 1
        
        If vsf.TextMatrix(lngLoop, mCol.结论描述) <> "" Then
            intCount = intCount + 1
            strTmp = strTmp & intCount & "、" & vsf.TextMatrix(lngLoop, mCol.结论描述) & vbCrLf
        End If
        
    Next
    strTmp = strTmp & vbCrLf
    
    '转储建议内容
    strTmp = strTmp & "二、建议：" & vbCrLf
    strTmp = strTmp & rtb.Text
    
    Text = strTmp
    
    Exit Property
    
errHand:
    
End Property

Public Sub ClearData()
    '------------------------------------------------------------------------------------------------------------------
    '功能:接口
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = ""
    
    rtb.Text = ""
End Sub

Public Function SaveData(lng病人ID As Long, lng主页ID As Long, lng病历ID As Long, strReturnSQL As String, strError As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:接口
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSQL() As String
    Dim strDate As String
    Dim LngCount As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(0 To vsf.Rows + 1)
    
    For lngLoop = 1 To vsf.Rows - 1
        If StrIsValid(vsf.TextMatrix(lngLoop, 1), 100) = False Then
            vsf.Row = lngLoop
            vsf.Col = 1
            vsf.ShowCell vsf.Row, vsf.Col
            Exit Function
        End If
    Next
    
    If StrIsValid(rtb.Text, 4000) = False Then
        rtb.SetFocus
    End If
    
    If chk(0).Value = 1 Then
        strDate = Format(msk.Text, "yyyy-MM-dd")
        If IsDate(strDate) = False Then
            strDate = ""
        Else
            strDate = strDate & " 00:00:00"
        End If
    End If
    
    strSQL(0) = "ZL_体检人员结论_DELETE(" & lng病历ID & ")"
    LngCount = 0
    
    For lngLoop = 1 To vsf.Rows - 1
        If Trim(vsf.TextMatrix(lngLoop, mCol.结论描述)) <> "" Then
            
            LngCount = LngCount + 1
            
            strSQL(lngLoop) = "ZL_体检人员结论_INSERT(" & lng病人ID & "," & _
                                                        lng主页ID & "," & _
                                                        lng病历ID & "," & _
                                                        "0," & _
                                                        LngCount & ",'" & _
                                                        vsf.TextMatrix(lngLoop, mCol.结论描述) & "','" & _
                                                        vsf.TextMatrix(lngLoop, mCol.异常结果) & "'," & _
                                                        "NULL," & _
                                                        Val(vsf.RowData(lngLoop)) & ",NULL,NULL," & _
                                                        Abs(Val(vsf.TextMatrix(lngLoop, mCol.疾病))) & "," & _
                                                        "'" & vsf.TextMatrix(lngLoop, mCol.诊断建议) & "')"
        End If
    Next
    strSQL(lngLoop + 1) = "ZL_体检人员结论_INSERT(" & lng病人ID & "," & _
                                                        lng主页ID & "," & _
                                                        lng病历ID & "," & _
                                                        "1," & _
                                                        "1," & _
                                                        "NULL,NULL,'" & _
                                                        rtb.Text & "'," & _
                                                        "NULL," & _
                                                        IIf(strDate = "", "Null", "To_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
                                                        IIf(chk(1).Value = 0, "Null", Val(txt.Text)) & "," & _
                                                        "0," & _
                                                        "NULL)"
        
    strTmp = ""
    For lngLoop = 0 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then
        
            strSQL(lngLoop) = Replace(strSQL(lngLoop), Chr(9), Chr(32))
            
            If strTmp = "" Then
                strTmp = strSQL(lngLoop)
            Else
                strTmp = strTmp & Chr(9) & strSQL(lngLoop)
            End If
        End If
    Next
    
    Dim strCallBack As String
    
    If chk(0).Value = 1 Then
        strCallBack = ""
        If Not (mrsCallBack Is Nothing) Then
            If mrsCallBack.State = adStateOpen Then
                If mrsCallBack.RecordCount > 0 Then
                    mrsCallBack.MoveFirst
                    Do While Not mrsCallBack.EOF
                        strCallBack = strCallBack & "," & Val(mrsCallBack("清单id").Value)
                        mrsCallBack.MoveNext
                    Loop
                    
                    If strCallBack <> "" Then strCallBack = Mid(strCallBack, 2)
                    strTmp = strTmp & Chr(9) & "ZL_体检登记记录_复查('" & mstr挂号单 & "'," & mlng病人id & ",To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strCallBack & "')"
                End If
            End If
            
        End If
    Else
        strTmp = strTmp & Chr(9) & "ZL_体检登记记录_复查('" & mstr挂号单 & "'," & mlng病人id & ",Null,Null)"
    End If
    
    '返回SQL语句
    strReturnSQL = strTmp
    
    SaveData = True
    
    Exit Function
    
errHand:

    strError = "体检专用纸保存失败！"
    
End Function

Private Function ShowOpenList(Optional strText As String, Optional ByVal bytMode As Byte = 1) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开列表结构的诊疗检验标本数据
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim strTitle As String
    Dim strDescrible As String
    
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    ShowOpenList = 2
    
    strText = "'%" & UCase(strText) & "%'"
    
    If bytMode = 1 Then
        
        strLvw = "编码,900,0,1;名称,1800,0,0;诊断建议,2700,0,0"
        strTitle = "体检结论过滤"
        strDescrible = "请从下表中选择一个体检结论"
        
        strSQL = _
                    "SELECT A.序号 AS ID, " & _
                            "A.编码, " & _
                            "A.名称, " & _
                            "A.是否疾病,A.诊断建议 " & _
                    "FROM 体检诊断建议 A " & _
                    "WHERE NVL(末级,0)=1 "
        strSQL = strSQL & " AND (A.编码 Like " & strText & " OR A.名称 Like " & strText & " OR A.简码 Like " & UCase(strText) & ")"
    End If
    
    Call OpenRecord(rs, strSQL, "体检结论")
    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then GoTo Over
    Call CalcPosition(sglX, sglY, vsf)
    If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 9000, 5100, strTitle, strDescrible) Then
        GoTo Over
    End If
    
    Exit Function
    
Over:
    If bytMode = 1 Then
        vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
        vsf.EditText = zlCommFun.NVL(rs("名称").Value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        
        vsf.TextMatrix(vsf.Row, mCol.疾病) = zlCommFun.NVL(rs("是否疾病").Value)
        vsf.TextMatrix(vsf.Row, mCol.诊断建议) = zlCommFun.NVL(rs("诊断建议").Value)
        
        Call ReadPreVisiteDate(2)
        
    End If
    
    Modified = True
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function ShowOpenTree(Optional ByVal bytMode As Byte = 1) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开列表结构的诊疗检验标本数据
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim strTitle As String
    Dim strDescrible As String
    
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    ShowOpenTree = 2
    
    If bytMode = 1 Then
        strLvw = "编码,900,0,1;名称,1800,0,0;诊断建议,2700,0,0"
        strTitle = "体检结论选择"
        strDescrible = "请从下表中选择一个体检诊断"
        
        strSQL = "SELECT -1 AS ID," & _
                            "0 AS 上级ID," & _
                            "0 AS 末级," & _
                            "'' AS 编码," & _
                            "'所有分类' AS 名称, " & _
                            "Null+0 AS 是否疾病,'' As 诊断建议 " & _
                    "FROM dual "
                    
        strSQL = strSQL & _
                " UNION ALL " & _
                "SELECT 序号 AS ID," & _
                            "DECODE(上级序号,NULL,-1,上级序号) AS 上级ID," & _
                            "0 AS 末级," & _
                            "编码," & _
                            "名称, " & _
                            "Null+0 AS 是否疾病,'' As 诊断建议 " & _
                    "FROM 体检诊断建议 " & _
                    "WHERE NVL(末级,0)=0 " & _
                    "START WITH 上级序号 is NULL CONNECT BY PRIOR 序号 = 上级序号 "
        
        strSQL = strSQL & _
                    "UNION ALL " & _
                    "SELECT A.序号 AS ID, " & _
                            "DECODE(上级序号,NULL,-1,上级序号) AS 上级ID, " & _
                            "1 AS 末级, " & _
                            "A.编码, " & _
                            "A.名称, " & _
                            "A.是否疾病,A.诊断建议 " & _
                    "FROM 体检诊断建议 A " & _
                    "WHERE NVL(A.末级,0)=1"
    End If
    
    Call OpenRecord(rs, strSQL, "体检结论")
    
    If rs.BOF Then
        ShowOpenTree = 0
        Exit Function
    End If
        
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectTree.ShowSelect(Screen, rs, sglX, sglY, 9000, 5100, vsf.CellHeight, strTitle, strLvw, strDescrible) Then
        GoTo Over
    End If
    
    Exit Function
    
Over:
    If bytMode = 1 Then
    
        vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
        vsf.EditText = zlCommFun.NVL(rs("名称").Value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        
        vsf.TextMatrix(vsf.Row, mCol.疾病) = zlCommFun.NVL(rs("是否疾病").Value, 0)
        vsf.TextMatrix(vsf.Row, mCol.诊断建议) = zlCommFun.NVL(rs("诊断建议").Value)
        
        Call ReadPreVisiteDate(2)
    End If
    
    Modified = True
    
    ShowOpenTree = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '------------------------------------------------------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft - 45
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight - 30
End Sub

Private Function GetAdvice() As String
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strSQL As String
        
    On Error GoTo errHand
    
    GetAdvice = ""
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            strSQL = "SELECT 参考建议 FROM 体检诊断建议 WHERE 序号 = " & Val(vsf.RowData(lngLoop))
            Call OpenRecord(rs, strSQL, "体检结论")
            If rs.BOF = False Then
                
                If zlCommFun.NVL(rs("参考建议").Value) <> "" Then
                    If vsf.TextMatrix(lngLoop, mCol.异常结果) <> "" Then
                        GetAdvice = GetAdvice & lngLoop & "." & vsf.TextMatrix(lngLoop, mCol.结论描述) & " {" & vsf.TextMatrix(lngLoop, mCol.异常结果) & "}：" & vbCrLf
                    Else
                        GetAdvice = GetAdvice & lngLoop & "." & vsf.TextMatrix(lngLoop, mCol.结论描述) & "：" & vbCrLf
                    End If
                    GetAdvice = GetAdvice & zlCommFun.NVL(rs("参考建议").Value) & vbCrLf & vbCrLf
                End If
                
            End If
            
        End If
    Next
    
    Exit Function
    
errHand:
        
End Function

Private Sub SetErr(lngErrNum As Long, strErr As String)
    '------------------------------------------------------------------------------------------------------------------
    '设置错误描述及错误号
    '如果lngErrNum=-1 表示 控件自己定义的错误
    '------------------------------------------------------------------------------------------------------------------
    
    mReturnErrnumber = lngErrNum
    mReturnErrDescription = strErr
End Sub

Private Function InDesign() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：判断当前运行程序是否在VB的工程环境中
    '------------------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
    
End Function

Private Sub ShowUsrControl(lngKey As Long, Optional ByVal blnEditMode As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------
    '功能：外部调用显示
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim intRow As Integer
    Dim blnSave As Boolean
    
    On Error GoTo errHand
    
    blnSave = Modified
    
    mDispMode = Not blnEditMode
    
    'Begin  <初始化处理>
    
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Sub
    
    'End    <初始化处理>


    'Begin  <读取数据>
    
    Call InitControl
    vsf.ExtendLastCol = True
    
    intRow = 0
    
    strSQL = "SELECT DISTINCT A.记录性质, A.记录序号,A.结论描述,A.异常结果,A.参考建议,A.结论id,A.是否疾病,A.诊断建议 FROM 体检人员结论 A WHERE A.病历id=" & mlng病历id & " ORDER BY A.记录性质,A.记录序号"
    Call OpenRecord(rs, strSQL, "体检专用纸")
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If zlCommFun.NVL(rs("记录性质").Value) = 0 Then
                
                intRow = intRow + 1
                vsf.Rows = intRow + 1
                
                vsf.RowData(intRow) = zlCommFun.NVL(rs("结论id").Value)
'                vsf.TextMatrix(intRow, 0) = zlCommFun.Nvl(rs("记录序号").Value) & "、"
                vsf.TextMatrix(intRow, mCol.结论描述) = zlCommFun.NVL(rs("结论描述").Value)
                vsf.TextMatrix(intRow, mCol.异常结果) = zlCommFun.NVL(rs("异常结果").Value)
                vsf.TextMatrix(intRow, mCol.疾病) = zlCommFun.NVL(rs("是否疾病").Value, 0)
                vsf.TextMatrix(intRow, mCol.诊断建议) = zlCommFun.NVL(rs("诊断建议").Value)
                                
            Else
                rtb.Text = zlCommFun.NVL(rs("参考建议").Value)
            End If
            
            rs.MoveNext
        Loop
        
        strSQL = "Select a.复查时间,a.随访期限 From 体检人员档案 a,病人病历内容 b Where a.体检病历id=b.病历记录id and b.id=[1]"
        Set rs = OpenSQLRecord(strSQL, "体检专用纸", mlng病历id)
        If rs.BOF = False Then
            
            If zlCommFun.NVL(rs("复查时间")) <> "" Then
                msk.Text = Format(zlCommFun.NVL(rs("复查时间")), "yyyy-MM-dd")
                chk(0).Value = 1
            End If
            
            If zlCommFun.NVL(rs("随访期限"), 0) > 0 Then
                txt.Text = zlCommFun.NVL(rs("随访期限"), 0)
                chk(1).Value = 1
            End If
        End If
    Else
        Call ReadPreVisiteDate
    End If
    
    'End    <读取数据>
        
    Modified = blnSave
    
    Exit Sub
    
errHand:

    If Ambient.UserMode = False Or InDesign = False Then
        SetErr Err.Number, Err.Description
        Exit Sub
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function EditRefresh(ByVal objVsf As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim LngCount As Long
    
    On Error GoTo errHand
        
    If MsgBox("是否要替换原来的总检结论？", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
        vsf.Rows = 2
        vsf.RowData(1) = 0
        vsf.TextMatrix(1, 1) = ""
    End If
    
    For lngLoop = 1 To objVsf.Rows - 1
        If Val(objVsf.RowData(lngLoop)) > 0 Then
            If Abs(Val(objVsf.TextMatrix(lngLoop, 0))) = 1 Then
                
                '检查Val(objVsf.RowData(lngLoop))是否已经存在
                For LngCount = 0 To vsf.Rows - 1
                    If Trim(vsf.TextMatrix(LngCount, 1)) = Trim(objVsf.TextMatrix(lngLoop, 1)) Then
                        GoTo NextLoop
                    End If
                Next
                
                If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then vsf.Rows = vsf.Rows + 1
                                
                vsf.RowData(vsf.Rows - 1) = Val(objVsf.RowData(lngLoop))
'                vsf.TextMatrix(vsf.Rows - 1, 0) = vsf.Rows - 1 & "、"
                vsf.TextMatrix(vsf.Rows - 1, 1) = objVsf.TextMatrix(lngLoop, 1)
                vsf.TextMatrix(vsf.Rows - 1, 2) = objVsf.TextMatrix(lngLoop, 2)
                vsf.TextMatrix(vsf.Rows - 1, 3) = Abs(Val(objVsf.Cell(flexcpData, lngLoop, 1, lngLoop, 1)))
                
            End If
        End If
        
NextLoop:
        
    Next
    
    Call ReadPreVisiteDate(2)
    
    EditRefresh = True
    
    Exit Function
    
errHand:
    
End Function

Private Function ReadPreVisiteDate(Optional ByVal bytMode As Byte = 1) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strAgain As String
    Dim lngVisit As Long
    
    On Error GoTo errHand
    
    Select Case bytMode
    Case 1      '从预约登记时的随访标志
    
        strSQL = "Select 随访期限 From 体检登记记录 Where 体检号=[1] and 随访期限 Is Not Null"
        Set rs = OpenSQLRecord(strSQL, "体检结论", mstr挂号单)
        If rs.BOF = False Then
            
            txt.Text = zlCommFun.NVL(rs("随访期限"))
                        
        End If
    Case 2          '从当前的结论中找取最大的期限及复查时间
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                strSQL = "Select 随访期限,复查间隔*30+sysdate As 复查时间 From 体检诊断建议 Where 序号=[1]"
                Set rs = OpenSQLRecord(strSQL, "体检结论", Val(vsf.RowData(lngLoop)))
                If rs.BOF = False Then
                    
                    If lngVisit < zlCommFun.NVL(rs("随访期限"), 0) Then lngVisit = zlCommFun.NVL(rs("随访期限"), 0)
                    If strAgain < Format(zlCommFun.NVL(rs("复查时间")), "yyyy-MM-dd") Then strAgain = Format(zlCommFun.NVL(rs("复查时间")), "yyyy-MM-dd")
                    
                End If
            End If
        Next
        
        If strAgain <> "" Then
            msk.Text = strAgain
        Else
            msk.Text = "____-__-__"
        End If
        txt.Text = lngVisit
    End Select
    
    chk(0).Value = IIf(msk.Text <> "" And msk.Text <> "____-__-__", 1, 0)
    chk(1).Value = IIf(Val(txt.Text) > 0, 1, 0)
    
    ReadPreVisiteDate = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ReadReportResult(Optional ByVal blnAdvice As Boolean = False) As Boolean
'------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long

    strSQL = "Select y.结论描述,y.异常结果,y.结论id,y.是否疾病,y.诊断建议,y.参考建议,y.复查时间,y.随访期限 " & _
                "From " & _
                "( " & _
                "Select b.医嘱id,d.排列顺序 " & _
                "From 病人医嘱记录 a,体检项目医嘱 b,体检项目清单 c,体检项目排列 d " & _
                "WHERE 病人来源=4 AND a.挂号单=[1] and a.病人id=[2] And a.相关id Is Null " & _
                      "and a.id=b.医嘱id " & _
                      "AND b.清单id=c.ID and c.诊疗项目id=d.诊疗项目id " & _
                ") x, " & _
                "( " & _
                "Select Distinct Nvl(a.相关id,a.id) As 医嘱id,d.病历id,d.结论描述,d.异常结果,d.结论id,d.是否疾病,d.诊断建议,d.记录性质,d.记录序号,d.参考建议,d.复查时间,d.随访期限 " & _
                "From 病人医嘱记录 a,病人医嘱发送 b,病人病历内容 c,体检人员结论 d " & _
                "WHERE a.病人来源=4 AND a.挂号单=[1] and a.病人id=[2] and d.记录性质=[3] and a.诊疗类别 In ('C','D') " & _
                      "and a.id=b.医嘱id and b.报告id Is Not Null and c.病历记录id=b.报告id and d.病历id=c.id " & _
                ") y " & _
                "where x.医嘱id(+)=y.医嘱id " & _
                "Order By  x.排列顺序,y.记录性质,y.记录序号"
                    
    If blnAdvice = False Then
        
        Set rs = OpenSQLRecord(strSQL, "体检结论", mstr挂号单, mlng病人id, 0)

        If rs.BOF = False Then
            Do While Not rs.EOF
                
                '如果没有,则填写
                If zlCommFun.NVL(rs("结论描述")) <> "" Then
                    vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("结论id"))
'                    vsf.TextMatrix(vsf.Rows - 1, 0) = vsf.Rows & "、"
                    vsf.TextMatrix(vsf.Rows - 1, mCol.结论描述) = zlCommFun.NVL(rs("结论描述"))
                    vsf.TextMatrix(vsf.Rows - 1, mCol.异常结果) = zlCommFun.NVL(rs("异常结果"))
                    vsf.TextMatrix(vsf.Rows - 1, mCol.疾病) = zlCommFun.NVL(rs("是否疾病"), 0)
                    vsf.TextMatrix(vsf.Rows - 1, mCol.诊断建议) = zlCommFun.NVL(rs("诊断建议"))
                    
                End If
                
                vsf.Rows = vsf.Rows + 1
                
                rs.MoveNext
            Loop
        End If
        
        If vsf.Rows > 1 Then vsf.Rows = vsf.Rows - 1
    
    Else
                                    
        Set rs = OpenSQLRecord(strSQL, "体检结论", mstr挂号单, mlng病人id, 1)
        If rs.BOF = False Then
            
            If zlCommFun.NVL(rs("复查时间")) <> "" Then
                msk.Text = Format(zlCommFun.NVL(rs("复查时间")), "yyyy-MM-dd")
                chk(0).Value = 1
                cmd.Visible = True
            End If
            
            If zlCommFun.NVL(rs("随访期限"), 0) > 0 Then
                txt.Text = zlCommFun.NVL(rs("随访期限"), 0)
                chk(1).Value = 1
            End If
            
            Do While Not rs.EOF
                
                rtb.Text = rtb.Text & Trim(zlCommFun.NVL(rs("参考建议"))) & vbCrLf
                
                rs.MoveNext
            Loop
        End If
    End If
    
    ReadReportResult = True
    
    Exit Function
    
errHand:
    
End Function

Private Sub chk_Click(Index As Integer)
    msk.Visible = (chk(0).Value = 1)
    cmd.Visible = (chk(0).Value = 1)
    txt.Visible = (chk(1).Value = 1)
    
    If (msk.Text = "" Or msk.Text = "____-__-__") And msk.Visible Then
        msk.Text = Format(zlDatabase.Currentdate + 90, "yyyy-MM-dd")
    End If
    
    Modified = True
End Sub

Private Sub cmd_Click()
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT A.ID,B.ID AS 清单id,E.登记id," & _
                  "DECODE(A.类别, 'C', '检验', 'D', '检查') AS 类别," & _
                  "A.名称," & _
                  "D.名称 as 执行科室," & _
                  "B.基本价格,"
    strSQL = strSQL & _
                  "E.组别名称, " & _
                  "B.采集方式id, " & _
                  "B.采集科室id, " & _
                  "B.执行科室id, " & _
                  "B.检查部位, " & _
                  "B.体检类型, " & _
                  "B.体检价格,Decode(b.基本价格,0,0,Null,0,10*B.体检价格/B.基本价格) As 折扣," & _
                  "B.检查部位id, " & _
                  "B.检验标本,Decode(F.复查清单id,0,0,Null,0,1) As 选择 " & _
             "FROM 诊疗项目目录 A,体检项目清单 B,部门表 D,体检人员档案 E,体检项目医嘱 F,体检登记记录 H " & _
            "WHERE B.执行科室id=D.ID(+) AND H.ID=E.登记id And A.ID = B.诊疗项目ID AND H.体检号=[1] AND E.登记id=B.登记id AND E.病人id=F.病人id AND F.清单id=B.ID AND F.病人id=[2] and F.复查清单id Is Null "
    
    strSQL = strSQL & " Order By A.名称"

    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "体检专用纸", mstr挂号单, mlng病人id)
    If ShowFilterDiagBox(Me, cmd, "名称,2700,0,0;类别,900,0,1;执行科室,1500,0,0", "体检专用纸\复查项目选择", "请从列表中选择要复查的体检项目。", rsData, mrsCallBack, 8790, 4500, , , True) Then
        
    End If
        
End Sub

Private Sub msk_Change()
    Modified = True
End Sub

Private Sub msk_GotFocus()
    zlControl.TxtSelAll msk
End Sub

Private Sub rtb_Change()
    Modified = True
End Sub

Private Sub cbr_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    Dim lng病人ID As Long
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    Select Case Index
    Case 0
        Select Case Button.Key
        Case "收集"
            
            If MsgBox("真的要从体检项目报告中提取结论吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            
            If MsgBox("是否要替换原来的总检结论？", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                vsf.Rows = 2
                vsf.RowData(1) = 0
                vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = ""
            End If
                        
            Call ReadReportResult(False)
            Call ReadPreVisiteDate(2)
            
        Case "评估"
            
            Call frmMedicalResult.ShowEdit(Me, mlng病人id & "'0'" & mstr挂号单)
            
        Case "清空"
            
            vsf.Rows = 2
            vsf.RowData(1) = 0
            vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = ""
            
            Modified = True
        End Select
    Case 1
    
        Select Case Button.Key
        Case "收集"
            If MsgBox("真的要从体检项目报告中提取建议吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            
            If MsgBox("是否要替换原来的总检建议？", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                rtb.Text = ""
            End If
                        
            Call ReadReportResult(True)
            
        Case "生成"
            rtb.Text = GetAdvice
        Case "清空"
            rtb.Text = ""
        End Select
        
    End Select
End Sub


Private Sub rtb_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub rtb_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt_Change()
    Modified = True
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub


Private Sub UserControl_Initialize()
    
    Call InitControl
    
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
        
    With fra(0)
        .Left = 0
        .Top = -90
        .Width = UserControl.Width
    End With
    
    With fra(1)
        .Left = 0
        .Top = fra(0).Top + fra(0).Height - 90
        .Width = fra(0).Width
        .Height = UserControl.Height + 90 - fraOther.Height + 90 - fra(0).Height - 90
    End With
    
    With fraOther
        .Left = fra(1).Left
        .Top = fra(1).Top + fra(1).Height - 90
        .Width = fra(1).Width
    End With
    
    With pic(0)
        .Left = 30
        .Top = 120
        .Width = fra(0).Width - .Left - 45
    End With

    
    With vsf
        .Left = 15
        .Top = pic(0).Top + pic(0).Height
        .Width = fra(0).Width - .Left - 30
    End With
       
    With pic(1)
        .Left = 30
        .Top = 120
        .Width = fra(1).Width - .Left - 45
    End With

                    
    With rtb
        .Left = 15
        .Top = pic(1).Top + pic(1).Height
        .Width = fra(1).Width - .Left - 30
        .Height = fra(1).Height - .Top - 30
    End With
    
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    
    Set mobjParentObject = Nothing
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    Call ReadPreVisiteDate(2)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Modified = True
    
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    If mDispMode Then Exit Sub
    
    Select Case Col
    Case mCol.结论描述
        
        Call ShowOpenTree(1)
        
    End Select
    
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim strSvrText As String
    
    If mDispMode Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        '对于2-文字型的情况
        
        If InStr(vsf.EditText, "'") > 0 Then
            KeyCode = 0
            Exit Sub
        End If

        strSvrText = vsf.EditText
        Select Case ShowOpenList(vsf.EditText)
        Case 0
            '没有匹配的项目
            vsf.Cell(flexcpData, Row, Col) = strSvrText
            
        Case 1
            '选取了一个项目
'            mblnChangeEdit = True
'            Call AdjustEnableState
        Case 2
            '取消了本次选择
            KeyCode = 0
            
            vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
            vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)

        End Select
    Else
'        mblnChangeEdit = True
'        Call AdjustEnableState
    End If
End Sub










