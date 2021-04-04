VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmModiPatiBaseInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人基本信息调整"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4965
   Icon            =   "frmModiPatiBaseInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton optType 
      Caption         =   "住院"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3390
      TabIndex        =   12
      Top             =   2085
      Width           =   870
   End
   Begin VB.OptionButton optType 
      Caption         =   "门诊"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2115
      TabIndex        =   11
      Top             =   2085
      Width           =   855
   End
   Begin VB.ComboBox cmbNum 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmModiPatiBaseInfo.frx":030A
      Left            =   2115
      List            =   "frmModiPatiBaseInfo.frx":030C
      TabIndex        =   14
      Text            =   "cmbNum"
      Top             =   2475
      Width           =   2070
   End
   Begin VB.ComboBox cboAge 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1590
      Width           =   705
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   2  'OFF
      Left            =   2115
      TabIndex        =   8
      Top             =   1590
      Width           =   1350
   End
   Begin VB.ComboBox cboSex 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmModiPatiBaseInfo.frx":030E
      Left            =   2115
      List            =   "frmModiPatiBaseInfo.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   675
      Width           =   2070
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2115
      MaxLength       =   100
      TabIndex        =   1
      Top             =   210
      Width           =   2070
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2010
      TabIndex        =   15
      Top             =   3210
      Width           =   1300
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3345
      TabIndex        =   16
      Top             =   3210
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   15
      TabIndex        =   17
      Top             =   2925
      Width           =   5310
   End
   Begin MSMask.MaskEdBox medBirthdayTime 
      Height          =   360
      Left            =   3480
      TabIndex        =   6
      Top             =   1140
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medBirthdayDate 
      Bindings        =   "frmModiPatiBaseInfo.frx":0312
      Height          =   360
      Left            =   2115
      TabIndex        =   5
      Top             =   1140
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   635
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "YYYY-MM-DD"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "就诊类别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1035
      TabIndex        =   10
      Top             =   2085
      Width           =   960
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "挂号单号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1035
      TabIndex        =   13
      Top             =   2535
      Width           =   960
   End
   Begin VB.Label lbl出生日期 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1035
      TabIndex        =   4
      Top             =   1200
      Width           =   960
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Left            =   495
      Picture         =   "frmModiPatiBaseInfo.frx":031D
      Top             =   345
      Width           =   480
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "年龄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1500
      TabIndex        =   7
      Top             =   1650
      Width           =   480
   End
   Begin VB.Label lblSex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1485
      TabIndex        =   2
      Top             =   750
      Width           =   480
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1530
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmModiPatiBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlng病人ID As Long
Private mlng就诊ID As Long
Private mstr模块 As String
Private mint场合 As Integer
Private mstrInfo As String
Private mblnChange As Boolean
Private mblnDrop As Boolean
Private mrsTmp As New ADODB.Recordset

Public Function ShowMe(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str模块 As String, ByRef strInfo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:lng病人ID-病人ID
    '     lng就诊ID=不等于0表示某一次住院的主页Id(程序将自动定位到要修改的某一次住院)，等于0表示需要用户手工选择是门诊还是住院
    '     str模块=调用该功能的模块描述，如"门诊挂号"，"检查报到"。
    '出参:
    '返回:
    '编制:刘鹏飞
    '日期:2013-10-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    mstr模块 = str模块
    mblnChange = False
    mblnOK = False
    '获取病人基本信息
    If Not LoadPatiBaseInfo Then ShowMe = False: Exit Function
    
    Me.Show 1, frmParent
    strInfo = Trim(mstrInfo)
    ShowMe = mblnOK
End Function

Private Sub InitDicts()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    txtName.Text = ""
    txtName.MaxLength = GetColumnLength("病人信息", "姓名")
    txtAge.Text = ""
    cboAge.Clear
    cboAge.AddItem "岁"
    cboAge.AddItem "月"
    cboAge.AddItem "天"
    cboAge.ListIndex = 0
    txtAge.MaxLength = GetColumnLength("病人信息", "年龄")
    
    cboSex.Clear
    
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "性别")
    Do While Not rsTmp.EOF
        cboSex.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If rsTmp!缺省 = 1 Then
            cboSex.ListIndex = cboSex.NewIndex
            cboSex.ItemData(cboSex.NewIndex) = 1
        End If
    rsTmp.MoveNext
    Loop
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadPatiBaseInfo() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngIndex As Long
    
    On Error GoTo ErrHand
    
    If mlng就诊ID <> 0 Then '住院病人
        strSQL = " Select Nvl(a.姓名, b.姓名) 姓名, Nvl(a.性别, b.性别) 性别,a.年龄,B.出生日期" & vbNewLine & _
                " From 病案主页 a, 病人信息 b" & vbNewLine & _
                " Where a.病人id = b.病人id And a.病人id = [1] And a.主页id = [2]"
    Else
        strSQL = "Select 姓名,性别,年龄,出生日期 From 病人信息 Where 病人ID=[1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病人基本信息", mlng病人ID, mlng就诊ID)
    
    mblnChange = False
    
    If Not rsTmp.EOF Then
        '基本信息初始化
        Call InitDicts
        
        txtName.Text = zlCommFun.Nvl(rsTmp!姓名)
        
        cboSex.ListIndex = GetCboIndex(cboSex, Nvl(rsTmp!性别))
        If cboSex.ListIndex = -1 And Not IsNull(rsTmp!性别) Then
            cboSex.AddItem rsTmp!性别, 0
            cboSex.ListIndex = cboSex.NewIndex
        End If
           
        Call LoadOldData("" & rsTmp!年龄, txtAge, cboAge)
        mblnChange = False
        medBirthdayDate.Text = Format(IIf(IsNull(rsTmp!出生日期), "____-__-__", rsTmp!出生日期), "YYYY-MM-DD")
        mblnChange = True
        
        If Not IsNull(rsTmp!出生日期) Then
            If CDate(medBirthdayDate.Text) - CDate(rsTmp!出生日期) <> 0 Then medBirthdayTime.Text = Format(rsTmp!出生日期, "HH:MM")
        Else
            medBirthdayTime.Text = "__:__"
            mblnChange = False
            medBirthdayDate.Text = ReCalcBirth(Val(txtAge.Text), cboAge.Text)
            mblnChange = True
        End If
    Else
        MsgBox "获取病人基本信息失败,请您确认要进行信息调整的病人！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call LoadPatiData
    
    mblnChange = True
    
    LoadPatiBaseInfo = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub LoadPatiData()
'-----------------------------------------------
'功能:提取病人就诊记录信息(住院次数或就诊记录)
'
'-----------------------------------------------
    Dim strSQL As String
    Dim bln门诊 As Boolean, bln住院 As Boolean
    
    On Error GoTo ErrHand
    strSQL = _
        " Select 1 性质,ID Id, No, to_char(登记时间,'YYYY-MM-DD hh24:mi:ss') 登记时间" & vbNewLine & _
        " From 病人挂号记录" & vbNewLine & _
        " Where 病人id = [1] And Mod(记录状态, 2) <> 0" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select 2 性质,主页Id Id, '' || 主页id No, to_char(登记时间,'YYYY-MM-DD hh24:mi:ss') 登记时间" & vbNewLine & _
        " From 病案主页" & vbNewLine & _
        " Where 病人id = [1] And Nvl(主页id, 0) <> 0"
    Set mrsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取就诊记录", mlng病人ID)
    
    optType(0).Enabled = True
    optType(1).Enabled = True
    cmbNum.Clear
    If mrsTmp.RecordCount > 0 Then
        mrsTmp.Filter = "性质=1"
        bln门诊 = mrsTmp.RecordCount > 0
        mrsTmp.Filter = "性质=2"
        bln住院 = mrsTmp.RecordCount > 0
        
        mblnChange = True
        If bln门诊 = True And bln住院 = True Then
            If mlng就诊ID <> 0 Then
                optType(1).Value = True
            Else
                optType(0).Value = True
            End If
        Else
            If bln门诊 = True Then
                optType(0).Value = True
                optType(1).Enabled = False
            Else
                optType(1).Value = True
                optType(0).Enabled = False
            End If
        End If
    Else
        mblnChange = False
        '病人从未挂号或住院
        optType(0).Value = True
        optType(0).Enabled = False
        optType(1).Enabled = False
        lblType.Enabled = False
        lblNum.Enabled = False
        cmbNum.Enabled = False
        mblnChange = True
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboAge_LostFocus()
    If Trim(txtAge.Text) = "" Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    If Not IsDate(medBirthdayDate.Text) Then
        mblnChange = False
        medBirthdayDate.Text = ReCalcBirth(Val(txtAge.Text), cboAge.Text)
        mblnChange = True
    End If
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboSex.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboSex.hwnd, KeyAscii)
    If lngIdx <> -2 Then cboSex.ListIndex = lngIdx
End Sub

Private Sub cmbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If cmbNum.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then
        mblnDrop = SendMessage(cmbNum.hwnd, &H157, 0, 0) = 1
    End If
End Sub

Private Sub cmbNum_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim rsTemp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        If cmbNum.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cmbNum.Text)
        If cmbNum.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cmbNum.List(cmbNum.ListIndex) Then Call zlControl.CboSetIndex(cmbNum.hwnd, -1)
        End If
        If strText = "" Then
            cmbNum.ListIndex = -1
        ElseIf cmbNum.ListIndex = -1 Then
            intIdx = -1
            strFilter = "性质=" & IIf(optType(0).Value = True, 1, 2)
            '先复制记录集
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrsTmp)
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrsTmp.Filter = strFilter: iCount = 0
            With mrsTmp
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrsTmp.EOF
                    Select Case intInputType
                    Case 0  '输入的是全数字
                        '如果输入的数字,需要检查:
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                        
                        
                        '主要是检查输入的内容与编号完全相同,则直接定位
                        If Nvl(!NO) = strText Then strResult = Nvl(!NO): iCount = 0: Exit Do
                        
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                        If Val(Nvl(!NO)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!NO)
                            iCount = iCount + 1
                        End If
                        
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                         If Val(Nvl(!NO)) Like strCompents Then
                            If isCheckExists(Nvl(!NO)) Then Call zlDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
                         End If
                    Case 1  '输入的是全字母
                        '规则:
                        ' 1.输入的简码相等,则直接定位
                        ' 2.根据参数来匹配相同数据
                        
                        '1.输入的简码相等,则直接定位
                        If Trim(Nvl(!NO)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!NO)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.根据参数来匹配相同数据
                        If Trim(Nvl(!NO)) Like strCompents Then
                            If isCheckExists(Nvl(!NO)) Then Call zlDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
                        End If
                    Case Else  ' 2-其他
                        '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                        '1.编码\简码相等,直接定位
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        
                        '1.编码\简码相等,直接定位
                        If Trim(!NO) = strText Then
                            If iCount = 0 Then strResult = Nvl(!NO)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        If Trim(!NO) Like strCompents Then
                            If isCheckExists(Nvl(!NO)) Then Call zlDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
                        End If
                    End Select
                    mrsTmp.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!NO)
            '直接定位
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheckExists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '需要检查是否有多条满足条件的记录
            If rsTemp.RecordCount <> 0 Then
                '先按某种方式进行排序
                If optType(0).Value = True Then
                    rsTemp.Sort = "登记时间 DESC"
                Else
                    rsTemp.Sort = "ID DESC"
                End If
                '弹出选择器
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1101, cmbNum, rsTemp, True, "", "性质", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '进行定位
                            If isCheckExists(Nvl(rsReturn!NO), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                '未找到
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cmbNum: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
             
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cmbNum.ListIndex = -1 Then
            cmbNum.Text = ""
            Exit Sub
        Else
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
            ElseIf intIdx <> cmbNum.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cmbNum.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If optType(0).Value = True Then
            If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then KeyAscii = 0
        Else
            If InStr("0123456789" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub cmbNum_Validate(Cancel As Boolean)
    If cmbNum.Text <> "" Then
        If GetCboIndex(cmbNum, NeedName(cmbNum.Text)) = -1 Then cmbNum.ListIndex = -1: cmbNum.Text = ""
    End If
    If cmbNum.Text = "" And cmbNum.Enabled = True Then '说明录入的信息，不存在列表中
        MsgBox "请选择" & IIf(optType(0).Value = True, "挂号单号", "住院次数"), vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Function isCheckExists(ByVal strNo As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在开单人下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '出参:
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cmbNum.ListCount - 1
        If NeedName(cmbNum.List(i)) = strNo Then
            If blnLocateItem Then cmbNum.ListIndex = i
            isCheckExists = True
            Exit Function
        End If
    Next
End Function


Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
'功能：完成数据校验和保存
    Dim strSQL As String, strInfo As String
    Dim str年龄 As String, str出生日期 As String, str性别 As String
    Dim lngTmp As Long
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    Dim D出生日期 As Date
    '第一步：数据合法性校验
    If Trim(txtName.Text) = "" Then
        MsgBox "必须输入病人的姓名！", vbInformation, gstrSysName
        If txtName.Enabled And txtName.Visible Then txtName.SetFocus: Exit Sub
    End If
    If cboSex.ListIndex = -1 Then
        MsgBox "必须确定病人的性别！", vbInformation, gstrSysName
        If cboSex.Enabled And cboSex.Visible Then cboSex.SetFocus: Exit Sub
    End If
    
    If Not IsDate(medBirthdayDate.Text) Then
        MsgBox "必须正确输入病人的出生日期！", vbInformation, gstrSysName
        If medBirthdayDate.Enabled And medBirthdayDate.Visible Then medBirthdayDate.SetFocus: Exit Sub
    End If
    
    If Trim(txtAge.Text) = "" Then
        MsgBox "必须输入病人的年龄！", vbInformation, gstrSysName
        If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus: Exit Sub
    End If
    
    If IsDate(medBirthdayDate.Text) Then
        lngTmp = GetOldAcademic(CDate(medBirthdayDate.Text), cboAge.Text)
        If (lngTmp <> 0 And lngTmp <> Val(txtAge.Text)) Or (lngTmp = 0 And lngTmp <> Val(txtAge.Text) And Not CDate(medBirthdayDate.Text) = CDate(0) And InStr(" 岁月天", cboAge.Text) > 1) Then
            strInfo = ""
            If lngTmp = 0 Then strInfo = ReCalcOld(CDate(medBirthdayDate.Text), cboAge, 0, False)
            If strInfo = "" Then
                strInfo = lngTmp & cboAge.Text
            End If
            If MsgBox("年龄和出生日期不一致，" & medBirthdayDate.Text & "出生现在应该是" & strInfo & "。" & _
                vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If Not CheckTextLength("姓名", txtName) Then Exit Sub
    If Not CheckTextLength("年龄", txtAge) Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    If cmbNum.Enabled And cmbNum.ListIndex = -1 Then
        MsgBox "必须选择" & IIf(optType(0).Value = True, "挂号单号", "住院次数") & "！", vbInformation, gstrSysName
        If cmbNum.Enabled And cmbNum.Visible Then cmbNum.SetFocus: Exit Sub
    End If
    
    
    If medBirthdayTime = "__:__" Then
        str出生日期 = IIf(IsDate(medBirthdayDate.Text), "TO_Date('" & medBirthdayDate.Text & "','YYYY-MM-DD')", "NULL")
        D出生日期 = CDate(Format(medBirthdayDate.Text, "YYYY-MM-DD"))
    Else
        str出生日期 = IIf(IsDate(medBirthdayDate.Text), "TO_Date('" & medBirthdayDate.Text & " " & medBirthdayTime.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
        D出生日期 = CDate(Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:mm:ss"))
    End If
    If InStr(1, cboSex.Text, "-") <> 0 Then
        str性别 = Split(cboSex.Text, "-")(1)
    Else
        str性别 = cboSex.Text
    End If
    
    str年龄 = Trim(txtAge.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cboAge.Text
    
    If cmbNum.Enabled = True Then
        mint场合 = IIf(optType(1).Value = True, 2, 1)
        mlng就诊ID = Val(cmbNum.ItemData(cmbNum.ListIndex))
    Else
        mint场合 = 1
        mlng就诊ID = 0
    End If
    
    '第二步：数据保存
     On Error GoTo ErrHand
    Set cmdTmp = New ADODB.Command
    strSQL = "Zl_病人信息_基本信息调整("
'   病人id_In 病人信息变动.病人id%Type,
    strSQL = strSQL & "" & mlng病人ID & ","
    Set cmdPara = cmdTmp.CreateParameter("病人ID", adVarNumeric, adParamInput, 18, mlng病人ID)
    cmdTmp.Parameters.Append cmdPara
'   就诊id_In Number := Null,
    strSQL = strSQL & "" & mlng就诊ID & ","
    Set cmdPara = cmdTmp.CreateParameter("就诊ID", adVarNumeric, adParamInput, 18, mlng就诊ID)
    cmdTmp.Parameters.Append cmdPara
'   模块_In   病人信息变动.变动模块%Type,
    strSQL = strSQL & "'" & mstr模块 & "',"
    Set cmdPara = cmdTmp.CreateParameter("变动模块", adVarChar, adParamInput, 100, mstr模块)
    cmdTmp.Parameters.Append cmdPara
'   姓名_In   病人信息.姓名%Type,
    strSQL = strSQL & "'" & Trim(txtName.Text) & "',"
    Set cmdPara = cmdTmp.CreateParameter("姓名", adVarChar, adParamInput, 100, Trim(txtName.Text))
    cmdTmp.Parameters.Append cmdPara
'   性别_In   病人信息.性别%Type,
    strSQL = strSQL & "'" & str性别 & "',"
    Set cmdPara = cmdTmp.CreateParameter("性别", adVarChar, adParamInput, 100, str性别)
    cmdTmp.Parameters.Append cmdPara
'   年龄_In   病人信息.年龄%Type
    strSQL = strSQL & "'" & str年龄 & "',"
    Set cmdPara = cmdTmp.CreateParameter("年龄", adVarChar, adParamInput, 100, str年龄)
    cmdTmp.Parameters.Append cmdPara
'   出生日期_In 病人信息.出生日期%Type,
    strSQL = strSQL & "" & str出生日期 & ","
'   场合_In   number(1)  --1-门诊;2-住院
    Set cmdPara = cmdTmp.CreateParameter("出生日期", adDBTimeStamp, adParamInput, , D出生日期)
    cmdTmp.Parameters.Append cmdPara
    strSQL = strSQL & "" & mint场合 & ","
    Set cmdPara = cmdTmp.CreateParameter("场合", adVarNumeric, adParamInput, 1, mint场合)
    cmdTmp.Parameters.Append cmdPara
'   说明_Out    Out 病人信息变动.说明%Type --出参
    strSQL = strSQL & "" & "" & ")"
    Set cmdPara = cmdTmp.CreateParameter("说明", adLongVarChar, adParamOutput, 4000)
    cmdTmp.Parameters.Append cmdPara
    cmdTmp.ActiveConnection = gcnOracle
    cmdTmp.CommandType = adCmdStoredProc
    cmdTmp.CommandText = "Zl_病人信息_基本信息调整"
    Call SQLTest(App.ProductName, "Zl_病人信息_基本信息调整", strSQL)
    cmdTmp.Execute
    Call SQLTest
    mstrInfo = Nvl(cmdTmp.Parameters("说明"), "")
    
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        If ActiveControl.Name <> txtName.Name And ActiveControl.Name <> txtAge.Name And ActiveControl.Name <> cmbNum.Name Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub medBirthdayDate_Change()
    If IsDate(medBirthdayDate.Text) And mblnChange Then
        mblnChange = False
        medBirthdayDate.Text = Format(CDate(medBirthdayDate.Text), "yyyy-mm-dd") '0002-02-02自动转换为2002-02-02,否则,看到的是2002,实际值却是0002
        mblnChange = True
        
        txtAge.Text = ReCalcOld(CDate(medBirthdayDate.Text), cboAge)
    End If
End Sub

Private Sub medBirthdayDate_GotFocus()
    Call OpenIme
    SelAll medBirthdayDate
End Sub

Private Sub medBirthdayDate_LostFocus()
    If medBirthdayDate.Text <> "____-__-__" And Not IsDate(medBirthdayDate.Text) Then
        medBirthdayDate.SetFocus
    End If
End Sub

Private Sub medBirthdayTime_GotFocus()
    Call OpenIme
    SelAll medBirthdayTime
End Sub

Private Sub medBirthdayTime_KeyPress(KeyAscii As Integer)
    If Not IsDate(medBirthdayDate) Then
        KeyAscii = 0
        medBirthdayTime.Text = "__:__"
    End If
End Sub

Private Sub medBirthdayTime_Validate(Cancel As Boolean)
    If medBirthdayTime.Text <> "__:__" And Not IsDate(medBirthdayTime.Text) Then
        medBirthdayTime.SetFocus
        Cancel = True
    End If
End Sub

Private Sub optType_Click(Index As Integer)
    If mblnChange = False Or mrsTmp Is Nothing Then Exit Sub
    If mrsTmp.State = adStateClosed Then Exit Sub
     
    If Index = 0 Then
        lblNum.Caption = "挂号单号"
        mrsTmp.Filter = "性质=1"
    ElseIf Index = 1 Then
        lblNum.Caption = "住院次数"
        mrsTmp.Filter = "性质=2"
    End If
    If Index = 0 Or Index = 1 Then
        cmbNum.Clear
        Do While Not mrsTmp.EOF
            cmbNum.AddItem Nvl(mrsTmp!NO)
            cmbNum.ItemData(cmbNum.NewIndex) = Val(mrsTmp!ID)
            If Index = 1 And mlng就诊ID = Val(mrsTmp!ID) Then
                cmbNum.ListIndex = cmbNum.NewIndex
            End If
        mrsTmp.MoveNext
        Loop
        
        If cmbNum.ListIndex = -1 And cmbNum.ListCount > 0 Then cmbNum.ListIndex = 0
    End If
End Sub

Private Sub txtAge_GotFocus()
    Call zlCommFun.OpenIme
    zlControl.TxtSelAll txtAge
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboAge.Visible = False And IsNumeric(txtAge.Text) Then
            Call txtAge_Validate(False)
            Call cboAge.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txtAge.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAge_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAge.Text) And Trim(txtAge.Text) <> "" Then
        cboAge.ListIndex = -1: cboAge.Visible = False
    ElseIf cboAge.Visible = False Then
        cboAge.ListIndex = 0: cboAge.Visible = True
    End If
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            Call CheckInputLen(txtName, KeyAscii)
        End If
    Else
        If Trim(txtName.Text) = "" Then
            Exit Sub
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme
    txtName.Text = Trim(txtName.Text)
End Sub
