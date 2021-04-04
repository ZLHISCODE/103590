VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "zlIDKind.ocx"
Begin VB.Form frmModiPatiBaseInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人基本信息调整"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4965
   Icon            =   "frmModiPatiBaseInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtInfo 
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
      TabIndex        =   16
      Top             =   3000
      Width           =   2070
   End
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
   Begin VB.TextBox txtPatient 
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
      TabIndex        =   17
      Top             =   3690
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
      TabIndex        =   18
      Top             =   3690
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   0
      TabIndex        =   19
      Top             =   3480
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
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   360
      Left            =   1410
      TabIndex        =   20
      ToolTipText     =   "快捷键F4"
      Top             =   210
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   635
      Appearance      =   2
      IDKindStr       =   $"frmModiPatiBaseInfo.frx":031D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "宋体"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "修改原因"
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
      Left            =   1095
      TabIndex        =   15
      Top             =   3060
      Width           =   960
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
      Left            =   1095
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
      Left            =   1095
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
      Left            =   1095
      TabIndex        =   4
      Top             =   1200
      Width           =   960
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Left            =   285
      Picture         =   "frmModiPatiBaseInfo.frx":03A4
      Top             =   375
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
      Left            =   1560
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
      Left            =   1545
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
      Left            =   885
      TabIndex        =   0
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "frmModiPatiBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mblnOK As Boolean
Private mlng病人ID As Long
Private mlng就诊ID As Long
Private mstr模块 As String
Private mint场合 As Integer
Private mstrAgeAndBirth As String     '记录修改前病人年龄和出生日期 格式："年龄_出生日期"
Private mblnChange As Boolean
Private mblnDrop As Boolean
Private mrsTmp As New ADODB.Recordset
Private mblnNotClick As Boolean
Private mblnBatch As Boolean
Private mstrName As String '记录调整病人基本信息前前病人的姓名

Public Function ShowMe(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal int场合 As Integer, ByVal str模块 As String, Optional ByVal blnBatch As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:lng病人ID-病人ID
    '     lng就诊ID=非0:挂号ID或主页ID(程序将自动定位到要修改的某一次住院或就诊)，等于0表示需要用户手工选择是门诊还是住院
    '     int场合 1-门诊;2-住院
    '     str模块=调用该功能的模块描述，如"门诊挂号"，"检查报到"。
    '出参:strInfo:信息调整导致的变化信息
    '返回:
    '编制:刘鹏飞
    '日期:2013-10-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    mint场合 = int场合
    mstr模块 = str模块
    If blnBatch = False Then
        If lng病人ID = 0 Or lng就诊ID = 0 Then
            MsgBox "非连续性调整病人信息时,需传入具体的病人ID和就诊ID!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    mblnBatch = blnBatch
    mblnChange = False
    mblnOK = False
    mblnNotClick = False
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub InitDicts()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    txtPatient.Text = ""
    txtPatient.MaxLength = gobjComlib.Sys.FieldsLength("病人信息", "姓名")
    txtAge.Text = ""
    cboAge.Clear
    cboAge.AddItem "岁"
    cboAge.AddItem "月"
    cboAge.AddItem "天"
    cboAge.ListIndex = 0
    txtAge.MaxLength = gobjComlib.Sys.FieldsLength("病人信息", "年龄")
    
    cboSex.Clear
    
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Order by 编码"
    Call gobjDatabase.OpenRecordset(rsTmp, strSQL, "性别")
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
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Function LoadPatiBaseInfo() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngIndex As Long
    
    On Error GoTo ErrHand
    
    Call ClearInfo
    
    If mlng就诊ID <> 0 Then
        If mint场合 = 1 Then '门诊病人
            strSQL = "Select  Nvl(a.姓名, b.姓名) 姓名, Nvl(a.性别, b.性别) 性别,nvl(a.年龄,b.年龄) 年龄,b.出生日期,B.病人类型,B.险类" & vbNewLine & _
                " From 病人挂号记录 A,病人信息 b" & vbNewLine & _
                " Where a.病人id = [1] And A.id=[2] and a.病人ID=B.病人ID And b.停用时间 is NULL"
        Else '住院病人
            strSQL = " Select Nvl(a.姓名, b.姓名) 姓名, Nvl(a.性别, b.性别) 性别,nvl(a.年龄,b.年龄) 年龄,B.出生日期,B.病人类型,B.险类,A.入院日期 " & vbNewLine & _
                    " From 病案主页 a, 病人信息 b" & vbNewLine & _
                    " Where a.病人id = b.病人id And a.病人id = [1] And a.主页id = [2] And b.停用时间 is NULL"
        End If
    Else
        strSQL = "Select 姓名,性别,年龄,出生日期,病人类型,险类 From 病人信息 Where 病人ID=[1] And 停用时间 is NULL"
    End If
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "提取病人基本信息", mlng病人ID, mlng就诊ID)
    
    mblnChange = False
    
    If Not rsTmp.EOF Then
        txtPatient.Text = gobjCommFun.Nvl(rsTmp!姓名)
        mstrName = gobjCommFun.Nvl(rsTmp!姓名)
        txtPatient.ForeColor = GetPatiColor(Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类) = True, &H80000008, vbRed))
        lblName.Tag = txtPatient.ForeColor
        cboSex.ListIndex = gobjComlib.cbo.FindIndex(cboSex, Nvl(rsTmp!性别), True)
        If cboSex.ListIndex = -1 And Not IsNull(rsTmp!性别) Then
            cboSex.AddItem rsTmp!性别, 0
            cboSex.ListIndex = cboSex.NewIndex
        End If
        Call gobjComlib.zlControl.LoadOldData("" & rsTmp!年龄, txtAge, cboAge)
        mblnChange = False
        medBirthdayDate.Text = Format(IIf(IsNull(rsTmp!出生日期), "____-__-__", rsTmp!出生日期), "YYYY-MM-DD")
        If Nvl(rsTmp!年龄) Like "约*" Or Trim(Nvl(rsTmp!年龄)) = "不详" Then
            If "" & rsTmp!出生日期 = "____-__-__" Then
                medBirthdayDate.Enabled = False
                medBirthdayTime.Enabled = False
            End If
        Else
            medBirthdayDate.Enabled = True
            medBirthdayTime.Enabled = True
        End If
        mblnChange = True
        If mlng就诊ID <> 0 And mint场合 = 2 Then medBirthdayDate.Tag = rsTmp!入院日期 & ""
        If Not IsNull(rsTmp!出生日期) Then
            If CDate(medBirthdayDate.Text) - CDate(rsTmp!出生日期) <> 0 Then
                mblnChange = False
                medBirthdayTime.Text = Format(rsTmp!出生日期, "HH:MM")
                mblnChange = True
            End If
        Else
            medBirthdayTime.Text = "__:__"
            mblnChange = False
            Call RecalcBirthDay
            mblnChange = True
        End If
    Else
        MsgBox "获取病人基本信息失败,请您重新确认要进行信息调整的病人！", vbInformation, gstrSysName
        mlng病人ID = 0: mlng就诊ID = 0
        If mblnBatch = True Then
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Else
            On Error Resume Next
            Unload Me
            Err.Clear
        End If
        Exit Function
    End If
    mstrAgeAndBirth = txtAge.Text & cboAge.Text & "_" & medBirthdayDate.Text & medBirthdayTime.Text
    Call LoadPatiData
    
    mblnChange = True
    
    LoadPatiBaseInfo = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub LoadPatiData()
'-----------------------------------------------
'功能:提取病人就诊记录信息(住院次数或就诊记录)
'
'-----------------------------------------------
    Dim strSQL As String
    Dim bln门诊 As Boolean, bln住院 As Boolean
    
    On Error GoTo ErrHand
    strSQL = "Select * From(" & _
        " Select 1 性质,ID Id, No,0 病人性质, to_char(登记时间,'YYYY-MM-DD hh24:mi:ss') 登记时间,NULL 病人类型,NULL 险类,NULL as 入院日期 " & vbNewLine & _
        " From 病人挂号记录" & vbNewLine & _
        " Where 病人id = [1] And Mod(记录状态, 2) <> 0" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select 2 性质,主页Id Id, '' || 主页id No,病人性质, to_char(登记时间,'YYYY-MM-DD hh24:mi:ss') 登记时间,病人类型,险类,入院日期 " & vbNewLine & _
        " From 病案主页" & vbNewLine & _
        " Where 病人id = [1] And Nvl(主页id, 0) <> 0) Order By No Desc"
    Set mrsTmp = gobjDatabase.OpenSQLRecord(strSQL, "读取就诊记录", mlng病人ID)
    
    optType(0).Enabled = True
    optType(1).Enabled = True
    cmbNum.Enabled = True
    cmbNum.Clear
    If mrsTmp.RecordCount > 0 Then
        mrsTmp.Filter = "性质=1"
        bln门诊 = mrsTmp.RecordCount > 0
        mrsTmp.Filter = "性质=2"
        bln住院 = mrsTmp.RecordCount > 0
        
        mblnChange = True
        If bln门诊 = True And bln住院 = True Then
            If mlng就诊ID <> 0 Then
                If mint场合 = 1 Then
                    optType(0).Value = True
                Else
                    optType(1).Value = True
                End If
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
        Call optType_Click(IIf(optType(0).Value = True, 0, 1))
    Else
        mblnChange = False
        '病人从未挂号或住院
        optType(0).Value = True
        optType(0).Enabled = False
        optType(1).Enabled = False
        cmbNum.Enabled = False
        mblnChange = True
    End If
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboAge_LostFocus()
    If Trim(txtAge.Text) = "" Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    If Not IsDate(medBirthdayDate.Text) Then
        mblnChange = False
        Call RecalcBirthDay
        mblnChange = True
    End If
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboSex.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = gobjComlib.cbo.MatchIndex(cboSex.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cboSex.ListIndex = lngIdx
End Sub

Private Sub cmbNum_Click()
    Dim lngColor As Long
    lngColor = Val(lblName.Tag)
    medBirthdayDate.Tag = ""
    If optType(0).Value = True Then
        txtPatient.ForeColor = lngColor
    Else
        If mrsTmp Is Nothing Then Exit Sub
        If mrsTmp.State = adStateClosed Then Exit Sub
        If optType(1).Value = True And cmbNum.ListIndex <> -1 Then
            mrsTmp.Filter = "性质=2 And ID=" & Val(cmbNum.ItemData(cmbNum.ListIndex))
            If mrsTmp.RecordCount > 0 Then
                lngColor = GetPatiColor(Nvl(mrsTmp!病人类型), IIf(IsNull(mrsTmp!险类) = True, &H80000008, vbRed))
                medBirthdayDate.Tag = mrsTmp!入院日期 & ""
            End If
        End If
        txtPatient.ForeColor = lngColor
    End If
End Sub

Private Sub cmbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If cmbNum.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then
        mblnDrop = SendMessage(cmbNum.hWnd, &H157, 0, 0) = 1
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
            Call gobjCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cmbNum.Text)
        If cmbNum.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cmbNum.List(cmbNum.ListIndex) Then Call gobjControl.CboSetIndex(cmbNum.hWnd, -1)
        End If
        If strText = "" Then
            cmbNum.ListIndex = -1
        ElseIf cmbNum.ListIndex = -1 Then
            intIdx = -1
            strFilter = "性质=" & IIf(optType(0).Value = True, 1, 2)
            '先复制记录集
            Set rsTemp = gobjDatabase.zlCopyDataStructure(mrsTmp)
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf gobjCommFun.IsCharAlpha(strText) Then
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
                            If isCheckExists(Nvl(!NO)) Then Call gobjDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
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
                            If isCheckExists(Nvl(!NO)) Then Call gobjDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
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
                            If isCheckExists(Nvl(!NO)) Then Call gobjDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
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
                If isCheckExists(strResult, True) Then gobjCommFun.PressKey vbKeyTab
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
                If gobjDatabase.zlShowListSelect(Me, glngSys, 1101, cmbNum, rsTemp, True, "", "性质", rsReturn) Then
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
                KeyAscii = 0: gobjControl.TxtSelAll cmbNum: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
             
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call gobjCommFun.PressKey(vbKeyTab)
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
                Call gobjCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
            End If
        End If
        Call gobjCommFun.PressKey(vbKeyTab)
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
        If gobjComlib.cbo.FindIndex(cmbNum, gobjComlib.ZLStr.NeedName(cmbNum.Text), True) = -1 Then cmbNum.ListIndex = -1: cmbNum.Text = ""
    End If
    If cmbNum.Text = "" And cmbNum.Enabled = True And cmbNum.ListCount > 0 Then '说明录入的信息，不存在列表中
        MsgBox "请选择" & IIf(optType(0).Value = True, "挂号单号", "住院次数"), vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Function isCheckExists(ByVal strNO As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在开单人下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '出参:
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cmbNum.ListCount - 1
        If IIf(optType(0).Value = True, gobjComlib.ZLStr.NeedName(cmbNum.List(i)), Val(cmbNum.List(i))) = strNO Then
            If blnLocateItem Then cmbNum.ListIndex = i
            isCheckExists = True
            Exit Function
        End If
    Next
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'功能：完成数据校验和保存
    Dim strInfo As String
    Dim str年龄 As String, str出生日期 As String, str性别 As String
    Dim lngTmp As Long
    Dim blnTrue As Boolean
    Dim blnEMPI As Boolean
    
    '第一步：数据合法性校验
    If mlng病人ID = 0 Then
        MsgBox "请您先确定要调整的病人!", vbInformation, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPatient.Text) = "" Then
        MsgBox "必须输入病人的姓名！", vbInformation, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Sub
    End If
    If cboSex.ListIndex = -1 Then
        MsgBox "必须确定病人的性别！", vbInformation, gstrSysName
        If cboSex.Enabled And cboSex.Visible Then cboSex.SetFocus: Exit Sub
    End If
    
    If medBirthdayDate.Enabled Then
        If Not IsDate(medBirthdayDate.Text) Then
            MsgBox "必须正确输入病人的出生日期！", vbInformation, gstrSysName
            If medBirthdayDate.Enabled And medBirthdayDate.Visible Then medBirthdayDate.SetFocus: Exit Sub
        End If
    End If
    If Trim(txtAge.Text) = "" Then
        MsgBox "必须输入病人的年龄！", vbInformation, gstrSysName
        If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus: Exit Sub
    End If
    '103905 修改原因必填最大长度为100个字符
    If Not gobjControl.TxtCheckInput(txtInfo, "修改原因") Then Exit Sub
    '76409,刘鹏飞,2014-08-06,年龄合法性检查
    str年龄 = txtAge.Text
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cboAge.Text
    If IsDate(medBirthdayDate.Text) Then
        If medBirthdayTime.Text = "__:__" Then
            str出生日期 = Format(medBirthdayDate.Text, "YYYY-MM-DD")
        Else
            str出生日期 = Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:MM:SS")
        End If
        If mstrAgeAndBirth = txtAge.Text & cboAge.Text & "_" & medBirthdayDate.Text & medBirthdayTime.Text Then
            '97836 只修改姓名时不做强制修改限制
            blnTrue = CheckAge(str年龄)
        Else
            If mint场合 = 2 And IsDate(medBirthdayDate.Tag) Then
                blnTrue = CheckAge(str年龄, str出生日期, , medBirthdayDate.Tag)
            Else
                blnTrue = CheckAge(str年龄, str出生日期)
            End If
        End If
    Else
        blnTrue = CheckAge(str年龄)
    End If
    If blnTrue = False Then
        If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus: Exit Sub
    End If
    
    If Not gobjComlib.zlControl.TxtCheckInput(txtPatient, "姓名") Then Exit Sub
    If Not gobjComlib.zlControl.TxtCheckInput(txtAge, "年龄") Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    If cmbNum.Enabled And cmbNum.ListIndex = -1 Then
        MsgBox "必须选择" & IIf(optType(0).Value = True, "挂号单号", "住院次数") & "！", vbInformation, gstrSysName
        If cmbNum.Enabled And cmbNum.Visible Then cmbNum.SetFocus: Exit Sub
    End If
    
    If medBirthdayDate.Enabled Then
        If medBirthdayTime = "__:__" Then
            str出生日期 = Format(medBirthdayDate.Text, "YYYY-MM-DD")
        Else
            str出生日期 = Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:mm")
        End If
    End If
    
    If InStr(1, cboSex.Text, "-") <> 0 Then
        str性别 = Split(cboSex.Text, "-")(1)
    Else
        str性别 = cboSex.Text
    End If
    
    str年龄 = Trim(txtAge.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cboAge.Text
    If cmbNum.ListIndex >= 0 Then
        mint场合 = IIf(optType(1).Value = True, 2, 1)
        mlng就诊ID = Val(cmbNum.ItemData(cmbNum.ListIndex))
    Else
        mint场合 = 1
        mlng就诊ID = 0
    End If
    strInfo = Trim(txtInfo.Text)
    'EMPI检查
    blnEMPI = EMPI_LoadPati(Trim(txtPatient.Text), str性别, str出生日期)
    '第二步：数据保存
    On Error GoTo ErrHand
    If Trim(txtPatient.Text) <> Trim(mstrName) Then
        If MsgBox("你是否将病人姓名【" & mstrName & "】调整为【" & txtPatient.Text & "】,是否真的调整吗？", vbYesNo, gstrSysName) = vbYes Then
            If SaveBaseInfo(mlng病人ID, mlng就诊ID, Trim(txtPatient.Text), str性别, str年龄, str出生日期, mstr模块, mint场合, strInfo, True, blnEMPI) = False Then
                If strInfo <> "" Then
                    MsgBox strInfo, vbInformation, gstrSysName
                End If
                Exit Sub
            End If
        Else
            txtPatient.SetFocus
            txtPatient.SelStart = 0
            txtPatient.SelLength = Len(txtPatient.Text)
            Exit Sub
        End If
    Else
        If SaveBaseInfo(mlng病人ID, mlng就诊ID, Trim(txtPatient.Text), str性别, str年龄, str出生日期, mstr模块, mint场合, strInfo, True, blnEMPI) = False Then
            If strInfo <> "" Then
                MsgBox strInfo, vbInformation, gstrSysName
            End If
            Exit Sub
        End If
    End If
    If strInfo <> "" Then
        MsgBox strInfo, vbInformation, gstrSysName
    End If
    mblnOK = True
    If mblnBatch = False Then Unload Me: Exit Sub
    mlng病人ID = 0: mlng就诊ID = 0
    Call ClearInfo
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intIndex As Integer
    If KeyCode = vbKeyReturn Then
       If ActiveControl.Name <> txtPatient.Name And ActiveControl.Name <> txtAge.Name And ActiveControl.Name <> cmbNum.Name Then
           Call gobjCommFun.PressKey(vbKeyTab)
       End If
    ElseIf KeyCode = vbKeyF4 And mblnBatch = True Then
        If Shift = vbCtrlMask And IDKind.Enabled Then
            intIndex = IDKind.GetKindIndex("IC卡号")
            If intIndex < 0 Then Exit Sub
            IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '基本信息初始化
    Call InitDicts
    
    If mblnBatch = True Then
        Call CreateMobjCard
        Call CreateSquareCardObject(Me, 1101)
         '初始化
        Call IDKind.zlInit(Me, 100, 1101, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
        
        If Not gobjSquare.objSquareCard Is Nothing Then
            IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    Else
        IDKind.Visible = False
        lblName.Left = lblSex.Left
    End If
    
    If mlng病人ID <> 0 Then
        Call LoadPatiBaseInfo
    Else
        Call ClearInfo
        txtPatient.Text = ""
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXml As String
    If mblnBatch = False Then Exit Sub
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call txtPatient_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, 1101, lng卡类别ID, False, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnBatch = False Then Exit Sub
    Set gobjSquare.objCurCard = objCard
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And mblnNotClick = False Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    gobjComlib.zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If mblnBatch = False Then Exit Sub
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub medBirthdayTime_Change()
    Dim strBirthday As String
    If IsDate(medBirthdayTime.Text) And IsDate(medBirthdayDate.Text) And mblnChange Then
        strBirthday = Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:MM:SS")
        If mint场合 = 2 And IsDate(medBirthdayDate.Tag) Then
            txtAge.Text = ReCalcOld(CDate(strBirthday), cboAge, , , CDate(medBirthdayDate.Tag))
        Else
            txtAge.Text = ReCalcOld(CDate(strBirthday), cboAge)
        End If
    End If
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If mblnBatch = False Then Exit Sub
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, txtPatient.Text, False)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If mblnBatch = False Then Exit Sub
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, txtPatient.Text, False)
End Sub

Private Sub medBirthdayDate_Change()
    Dim strBirthday As String
    If IsDate(medBirthdayDate.Text) And mblnChange Then
        mblnChange = False
        medBirthdayDate.Text = Format(CDate(medBirthdayDate.Text), "yyyy-mm-dd") '0002-02-02自动转换为2002-02-02,否则,看到的是2002,实际值却是0002
        mblnChange = True
        If medBirthdayTime.Text = "__:__" Then
            strBirthday = Format(medBirthdayDate.Text, "YYYY-MM-DD")
        Else
            strBirthday = Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:MM:SS")
        End If
        If mint场合 = 2 And IsDate(medBirthdayDate.Tag) Then
            txtAge.Text = ReCalcOld(CDate(strBirthday), cboAge, , , CDate(medBirthdayDate.Tag))
        Else
            txtAge.Text = ReCalcOld(CDate(strBirthday), cboAge)
        End If
    End If
End Sub

Private Sub medBirthdayDate_GotFocus()
    Call gobjCommFun.OpenIme
    gobjComlib.zlControl.TxtSelAll medBirthdayDate
End Sub

Private Sub medBirthdayDate_LostFocus()
    If medBirthdayDate.Text <> "____-__-__" And Not IsDate(medBirthdayDate.Text) Then
        medBirthdayDate.SetFocus
    End If
End Sub

Private Sub medBirthdayTime_GotFocus()
    Call gobjCommFun.OpenIme
    gobjComlib.zlControl.TxtSelAll medBirthdayTime
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
            cmbNum.AddItem Nvl(mrsTmp!NO) & IIf(Val("" & mrsTmp!病人性质) = 1, "-门诊留观", IIf(Val("" & mrsTmp!病人性质) = 2, "-住院留观", ""))
            cmbNum.ItemData(cmbNum.NewIndex) = Val(mrsTmp!ID)
            If mlng就诊ID = Val(mrsTmp!ID) Then
                cmbNum.ListIndex = cmbNum.NewIndex
            End If
        mrsTmp.MoveNext
        Loop
        
        If cmbNum.ListIndex = -1 And cmbNum.ListCount > 0 Then cmbNum.ListIndex = 0
        cmbNum.Enabled = mblnBatch
    End If
End Sub

Private Sub txtAge_GotFocus()
    Call gobjCommFun.OpenIme
    gobjControl.TxtSelAll txtAge
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboAge.Visible = False And IsNumeric(txtAge.Text) Then
            Call txtAge_Validate(False)
            Call cboAge.SetFocus
        Else
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txtAge.Text) Then Call gobjCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAge_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAge.Text) And Trim(txtAge.Text) <> "" Then
        If Not Trim(txtAge.Text) Like "约*" And Trim(txtAge.Text) <> "不详" Then
            cboAge.ListIndex = -1: cboAge.Visible = False
            medBirthdayDate.Enabled = True
            medBirthdayTime.Enabled = True
        ElseIf Trim(txtAge.Text) Like "约*" Or Trim(txtAge.Text) = "不详" Then
            If Trim(medBirthdayDate.Text) = "____-__-__" Then
                medBirthdayDate.Enabled = False
                medBirthdayTime.Enabled = False
            End If
            cboAge.ListIndex = -1: cboAge.Visible = False
        End If
    ElseIf cboAge.Visible = False Or medBirthdayDate.Enabled = True Then
        cboAge.ListIndex = 0: cboAge.Visible = True
        medBirthdayDate.Enabled = True
        medBirthdayTime.Enabled = True
    Else
        medBirthdayDate.Enabled = True
        medBirthdayTime.Enabled = True
    End If
End Sub

Private Sub txtInfo_GotFocus()
    Call gobjCommFun.OpenIme
    gobjControl.TxtSelAll txtInfo
End Sub

Private Sub txtPatient_Change()
    If mblnBatch = False Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    gobjComlib.zlControl.TxtSelAll txtPatient
    If mblnBatch = False Then Exit Sub
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjIDCard.SetEnabled (True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjICCard.SetEnabled (True)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnBatch = False Then Exit Sub
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.ActiveFastKey
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim blnCard As Boolean, blnName As Boolean
    
    If Trim(txtPatient.Text) = "" Then
        Exit Sub
    End If
    
    If mblnBatch = False Then
        If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    
    If IDKind.GetCurCard.名称 = "门诊号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    End If
    
    If IDKind.GetCurCard.名称 Like "姓名*" Then
        blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("门诊号") Or IDKind.IDKind = IDKind.GetKindIndex("住院号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If
    
    '刷卡完毕或输入号码后回车
    If blnCard And Len(Me.txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        Call FindPati(IDKind.GetCurCard, txtPatient.Text, blnCard)
    End If
End Sub

Private Sub txtPatient_LostFocus()
    If mblnBatch = True Then
        If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
        If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    End If
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub CreateMobjCard()
    '创建卡部件
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
End Sub

Private Function FindPati(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
    '读取病人信息
    Dim blnName As Boolean
    If Not GetPatient(objCard, strInput, blnCard, blnName) Then
        If IsNumeric(txtPatient.Text) Then
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.Text = ""
        End If
        Call gobjControl.TxtSelAll(txtPatient)
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        If blnName = True Then gobjCommFun.PressKey vbKeyTab
    Else
        txtPatient.PasswordChar = ""
        txtPatient.IMEMode = 0
        Call LoadPatiBaseInfo
    End If
    
    FindPati = True
End Function

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False, Optional blnName As Boolean = False) As Boolean
'功能：读取病人信息
    Dim lng卡类别ID As Long, lng病人ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnDo As Boolean
    Dim blnHavePassWord As Boolean
    Dim strPassWord As String, strErrMsg As String
    Dim strCard As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select A.病人ID,A.姓名,A.性别,A.年龄,A.出生日期" & _
        " From 病人信息 A" & _
        " Where A.停用时间 is NULL"
        
    If blnCard = True And objCard.名称 Like "姓名*" Then    '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strSQL = strSQL & " And A.住院号=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " And A.门诊号=[1]"
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                blnName = (mlng病人ID > 0)
                Exit Function
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.医保号=[2]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.门诊号=[2]"
            Case Else
                '其他类别的,获取相关的病人ID
                If Val(objCard.接口序号) > 0 Then
                    lng卡类别ID = Val(objCard.接口序号)
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    blnDo = Not rsTmp.EOF
    
    If blnDo Then
        mlng病人ID = rsTmp!病人ID
        mlng就诊ID = 0
        GetPatient = True
    Else
NotFoundPati:
        mlng病人ID = 0
        mlng就诊ID = 0
        Call ClearInfo
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub ClearInfo()
    mblnChange = False
    mstrAgeAndBirth = ""
    Set mrsTmp = New ADODB.Recordset
    txtPatient.Tag = ""
    txtPatient.Text = ""
    txtPatient.ForeColor = &H80000008
    lblName.Tag = txtPatient.ForeColor
    medBirthdayDate.Text = "____-__-__"
    medBirthdayTime.Text = "__:__"
    medBirthdayDate.Tag = ""
    txtAge.Text = ""
    txtInfo.Text = ""
    cmbNum.Clear
    optType(0).Value = True
    optType(0).Enabled = False
    optType(1).Enabled = False
    cmbNum.Enabled = False
    mblnChange = True
    mstrName = ""
End Sub

Private Function EMPI_LoadPati(ByVal str姓名 As String, ByVal str性别 As String, ByVal str出生日期 As String) As Boolean
'功能:将EMPI返回来的病人信息更新到界面
    Dim rsPatiIn As ADODB.Recordset
    Dim rsPatiOut As ADODB.Recordset
    Dim blnRet As Boolean
    
    If CreatePlugInOK(glngModule) Then
        '组织病人基本信息
        Set rsPatiIn = New ADODB.Recordset
        With rsPatiIn.Fields
            .Append "病人ID", adBigInt
            .Append "主页ID", adBigInt
            .Append "挂号ID", adBigInt
            '-------------------------------
            .Append "门诊号", adVarChar, 18
            .Append "住院号", adVarChar, 18
            .Append "医保号", adVarChar, 30
            .Append "身份证号", adVarChar, 18
            .Append "其他证件", adVarChar, 20
            .Append "姓名", adVarChar, 100
            .Append "性别", adVarChar, 4
            .Append "出生日期", adVarChar, 20 '日期格式：YYYY-MM-DD HH:MM:SS
            .Append "出生地点", adVarChar, 100
            .Append "国籍", adVarChar, 30
            .Append "民族", adVarChar, 20
            .Append "学历", adVarChar, 10
            .Append "职业", adVarChar, 80
            .Append "工作单位", adVarChar, 100
            .Append "邮箱", adVarChar, 30
            .Append "婚姻状况", adVarChar, 4
            .Append "家庭电话", adVarChar, 20
            .Append "联系人电话", adVarChar, 20
            .Append "单位电话", adVarChar, 20
            .Append "家庭地址", adVarChar, 100
            .Append "家庭地址邮编", adVarChar, 6
            .Append "户口地址", adVarChar, 100
            .Append "户口地址邮编", adVarChar, 6
            .Append "单位邮编", adVarChar, 6
            .Append "联系人地址", adVarChar, 100
            .Append "联系人关系", adVarChar, 30
            .Append "联系人姓名", adVarChar, 64
        End With
        rsPatiIn.CursorLocation = adUseClient
        rsPatiIn.LockType = adLockOptimistic
        rsPatiIn.CursorType = adOpenStatic
        rsPatiIn.Open
         '1-门诊;2-住院(lng就诊ID=0,则默认为1;lng就诊ID<>0,1-lng就诊ID为挂号ID,2-lng就诊ID为主页ID)
        With rsPatiIn
            .AddNew
            !病人ID = mlng病人ID
            !主页ID = IIf(mlng就诊ID <> 0, IIf(mint场合 = 2, mlng就诊ID, 0), 0)
            !挂号ID = IIf(mlng就诊ID <> 0, IIf(mint场合 = 1, mlng就诊ID, 0), 0)
            !姓名 = str姓名
            !性别 = str性别
            !出生日期 = str出生日期
            .Update
            '-------------------------------------------------------
        End With
        
        '调用查询接口
        On Error Resume Next
        blnRet = gobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModule, rsPatiIn, rsPatiOut)
        If Err.Number = 438 Then blnRet = False
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: On Error GoTo 0
        If Not blnRet Then Exit Function
        If rsPatiOut Is Nothing Then Exit Function
        If rsPatiOut.RecordCount = 0 Then Exit Function
        EMPI_LoadPati = True      '用于标记找到建档病人
    End If
End Function

Private Sub RecalcBirthDay()
'功能:通过年龄反推出生日期
    Dim strBirth As String
    
    If RecalcBirth(Trim(txtAge.Text) & IIf(cboAge.Visible, Trim(cboAge.Text), ""), strBirth) Then
        If medBirthdayDate.Enabled Then medBirthdayDate.Text = Format(strBirth, "YYYY-MM-DD")
        If medBirthdayTime.Enabled Then medBirthdayTime.Text = IIf(Format(strBirth, "HH:MM") = "00:00", "__:__", Format(strBirth, "HH:MM"))
    End If
End Sub
