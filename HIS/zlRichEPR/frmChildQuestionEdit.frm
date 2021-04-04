VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmChildQuestionEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   4995
      Index           =   1
      Left            =   45
      ScaleHeight     =   4995
      ScaleWidth      =   5085
      TabIndex        =   22
      Top             =   480
      Width           =   5085
      Begin VB.Frame fra 
         Height          =   4935
         Left            =   390
         TabIndex        =   23
         Top             =   -15
         Width           =   4425
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   3
            ItemData        =   "frmChildQuestionEdit.frx":0000
            Left            =   810
            List            =   "frmChildQuestionEdit.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1275
            Width           =   1365
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   2
            ItemData        =   "frmChildQuestionEdit.frx":001E
            Left            =   795
            List            =   "frmChildQuestionEdit.frx":0028
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1590
            Width           =   1365
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   2745
            MaxLength       =   5
            TabIndex        =   26
            Top             =   1605
            Width           =   960
         End
         Begin VB.TextBox txt 
            Height          =   750
            Index           =   8
            Left            =   795
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   1935
            Width           =   2655
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   4
            Left            =   2745
            MaxLength       =   5
            TabIndex        =   7
            Top             =   1275
            Width           =   960
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   2010
            TabIndex        =   2
            Text            =   "cbo"
            Top             =   135
            Width           =   2205
         End
         Begin VB.TextBox txt 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   3
            Left            =   795
            TabIndex        =   21
            Top             =   4590
            Width           =   2655
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   795
            TabIndex        =   13
            Top             =   3045
            Width           =   2655
         End
         Begin VB.TextBox txt 
            BackColor       =   &H80000000&
            Height          =   510
            Index           =   1
            Left            =   795
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   3705
            Width           =   2655
         End
         Begin VB.TextBox txt 
            Height          =   750
            Index           =   0
            Left            =   795
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   5
            Left            =   795
            TabIndex        =   11
            Top             =   2715
            Width           =   2655
         End
         Begin VB.TextBox txt 
            BackColor       =   &H80000000&
            Height          =   300
            Index           =   6
            Left            =   795
            TabIndex        =   19
            Top             =   4260
            Width           =   2655
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   7
            Left            =   795
            TabIndex        =   15
            Top             =   3375
            Width           =   2655
         End
         Begin MSComctlLib.Toolbar tbrFree 
            Height          =   450
            Left            =   285
            TabIndex        =   24
            Top             =   720
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   794
            ButtonWidth     =   820
            ButtonHeight    =   794
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "imgColor24"
            DisabledImageList=   "img41"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "自由"
                  Object.ToolTipText     =   "自由录入反馈意见(F3)"
                  ImageIndex      =   1
                  Style           =   1
               EndProperty
            EndProperty
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            ItemData        =   "frmChildQuestionEdit.frx":0040
            Left            =   795
            List            =   "frmChildQuestionEdit.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   135
            Width           =   1215
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   0
            Left            =   3465
            Picture         =   "frmChildQuestionEdit.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   465
            Width           =   300
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "分制"
            Height          =   180
            Index           =   12
            Left            =   390
            TabIndex        =   30
            Top             =   1335
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "评分级别"
            Height          =   180
            Index           =   11
            Left            =   30
            TabIndex        =   28
            Top             =   1650
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "评分次"
            Height          =   180
            Index           =   10
            Left            =   2160
            TabIndex        =   25
            Top             =   1665
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "补充说明"
            Height          =   180
            Index           =   9
            Left            =   30
            TabIndex        =   8
            Top             =   1965
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "扣分数"
            Height          =   180
            Index           =   8
            Left            =   2160
            TabIndex        =   6
            ToolTipText     =   "分值"
            Top             =   1335
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "处理时间"
            Height          =   180
            Index           =   6
            Left            =   30
            TabIndex        =   20
            Top             =   4620
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "反馈时间"
            Height          =   180
            Index           =   5
            Left            =   30
            TabIndex        =   12
            Top             =   3075
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "处理说明"
            Height          =   180
            Index           =   4
            Left            =   30
            TabIndex        =   16
            Top             =   3720
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "处 理 人"
            Height          =   180
            Index           =   3
            Left            =   30
            TabIndex        =   18
            Top             =   4290
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "反 馈 人"
            Height          =   180
            Index           =   2
            Left            =   30
            TabIndex        =   10
            Top             =   2745
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "反馈意见"
            Height          =   180
            Index           =   1
            Left            =   30
            TabIndex        =   3
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "反馈对象"
            Height          =   180
            Index           =   0
            Left            =   30
            TabIndex        =   0
            Top             =   195
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "处理期限"
            Height          =   180
            Index           =   7
            Left            =   30
            TabIndex        =   14
            Top             =   3420
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.ImageList imgColor24 
      Left            =   3885
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChildQuestionEdit.frx":6896
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img41 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChildQuestionEdit.frx":6F90
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChildQuestionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain            As Object
Private mlngKey             As Long
Private mstr文件ID          As String
Private mlng医嘱id          As Long
Private mlng科室ID          As Long

Private mlngReferKey        As Long
Private mblnReading         As Boolean
Private mstrSQL             As String
Private mblnDataChanged     As Boolean
Private mblnAllowModify     As Boolean
Private mbytMode            As Byte
Private mstrObject          As String
Private mlng次数            As Long
Private mrsCondition        As ADODB.Recordset
Private mstrPrivs           As String
Private Type Items
    反馈意见                As String
End Type

Private mblnAuditEnter  As Boolean
Private mlng分值        As Long
Private mRsType As ADODB.Recordset
Private mrsEmr As ADODB.Recordset
Private mblnReadCom As Boolean

Private Type AudtiObject
    strName As String
    strID As String
    strPara As String
End Type

Private mTypeAuditObject() As AudtiObject
Private usrSaveItem As Items
Private zlCheck             As New clsCheck
Public Event AfterDataChanged()
Public Event AfterQuestionType(ByVal blnQuestionType As Boolean)
Public Event AfterParaments(ByVal strObject As String, ByVal strParam As String)
Public Event RefStatus()

Public Property Let AllowModify(blnData As Boolean)
    mblnAllowModify = blnData
    Call ExecuteCommand("控件状态")
End Property

Public Property Get AllowModify() As Boolean
    AllowModify = mblnAllowModify
End Property

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True, Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Set mfrmMain = frmMain
    mblnAllowModify = blnAllowModify
    mstrPrivs = strPrivs
    
    If ExecuteCommand("初始控件") = False Or ExecuteCommand("初始数据") = False Then Exit Function
    Call ExecuteCommand("控件状态")
        
    DataChanged = False
    
End Function

Public Function ClearData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    ClearData = ExecuteCommand("清空数据")
End Function

Public Function RefreshData(ByVal lngKey As Long, ByVal blnAuditEnter As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngReferKey = 0
    mlngKey = lngKey
    mbytMode = 2
    mblnAuditEnter = blnAuditEnter
    
    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("初始数据")
    

    If mlngKey > 0 Then
        If ExecuteCommand("读取数据", mlngKey) = False Then Exit Function
    End If
    Call ExecuteCommand("控件状态")
    
    DataChanged = False
    
    RefreshData = True
    
End Function

Public Sub SetCurNum(ByVal lngCur次数 As Long)
    mlng次数 = lngCur次数
    txt(9).Text = lngCur次数
End Sub

Public Function NewData(ByVal strObject As String, ByVal str文件id As String, ByVal lng医嘱id As Long, ByVal lng科室ID As Long, Optional ByVal lngReferKey As Long = 0, Optional ByVal lng次数 As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    mlngKey = 0
    mstr文件ID = str文件id
    mlng医嘱id = lng医嘱id
    mlng科室ID = lng科室ID
    mlng次数 = lng次数
    
    mlngReferKey = lngReferKey
    mstrObject = strObject
    
    mbytMode = 1
    
    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("初始数据")
    Call ExecuteCommand("控件状态")
    Call ExecuteCommand("读取数据", mlngReferKey)
    Call ExecuteCommand("缺省数据")
    
    DataChanged = True
    
    Call LocationObj(txt(0))
        
    NewData = True
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    If StrIsValid(txt(0).Text, txt(0).MaxLength) = False Then
        If txt(0).Enabled Then Call zlControl.TxtSelAll(txt(0)): txt(0).SetFocus
        Exit Function
    End If
    
    If StrIsValid(txt(8).Text, txt(8).MaxLength) = False Then
        If txt(8).Enabled Then Call zlControl.TxtSelAll(txt(8)): txt(8).SetFocus
        Exit Function
    End If
    
    If cbo(3).ItemData(cbo(3).ListIndex) = 0 Then
        '扣分制
        If Val(mlng分值) <> 0 Then
            If Val(txt(4).Text) > Val(mlng分值) Then
                If txt(4).Enabled Then Call zlControl.TxtSelAll(txt(4)): txt(4).SetFocus
                Exit Function
            End If
        End If
    End If
    
    ValidData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset, ByRef lngKey As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal lng提交Id As Long, Optional ByVal mlng次数 As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim strLimitDate As String
    Dim strFileID As String, strSubid As String
    Dim strPalValue As String '分制
    Dim strScore As String '分值
    Dim strPostscript As String '补充说明
    Dim strNum As String   '次数
    Dim strSelQuestion As String '反馈记录
    Dim strGradeRank   As String '评分级别
    Dim strTempValue As String
    
    
    On Error GoTo errHand
    
    If mlngKey = 0 Then
        '新增
        lngKey = zlDatabase.GetNextId("病案反馈记录")
    Else
        '修改
        lngKey = mlngKey
    End If
        
    strLimitDate = IIf(txt(7).Text = "", "Null", "To_Date('" & txt(7).Text & "','yyyy-mm-dd hh24:mi:ss')")
    strScore = IIf(txt(4).Text = "", "Null", txt(4).Text)
    strPostscript = txt(8).Text
    strPalValue = cbo(3).ItemData(cbo(3).ListIndex)
    strSelQuestion = cbo(1).Text
    strGradeRank = cbo(2).ItemData(cbo(2).ListIndex)
    
    If mlng次数 = 0 Then
        strNum = "Null"
    Else
        strNum = mlng次数
    End If
    
    If cbo(1).ListIndex >= 0 Then
        strFileID = cbo(1).ItemData(cbo(1).ListIndex)
        If strFileID = 0 Then
            strFileID = cbo(1).Tag
            If strFileID <> "" Then
                strSubid = Split(strFileID, "|")(1)
                strFileID = Split(strFileID, "|")(0)
            Else
                strFileID = 0
            End If
        End If
    End If
    
    If mlngReferKey = -1 Then
        strTempValue = "Null"
    Else
        strTempValue = mlngReferKey
    End If
    
    strSQL = "zl_病案反馈记录_Update(" & lngKey & "," & strTempValue & "," & lng提交Id & "," & lng病人ID & "," & lng主页ID & "," & cbo(0).ItemData(cbo(0).ListIndex) & ",'" & strFileID & "','" & txt(0).Text & "'," & Val(cmd(0).Tag) & ",'" & txt(5).Text & "',To_Date('" & txt(2).Text & "','yyyy-mm-dd hh24:mi:ss')," & strLimitDate & "," & mlng医嘱id & "," & mlng科室ID & "," & strGradeRank & "," & strPalValue & "," & strScore & ",'" & strPostscript & "'," & strNum & ",'" & strSelQuestion & "','" & strSubid & "')"
    Call SQLRecordAdd(rsSQL, strSQL)
            
    SaveData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'######################################################################################################################
Private Function ExecuteCommand(ByVal strCmd As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim blnAllowModify As Boolean
        
    On Error GoTo errHand
    
    mblnReading = True
    Call SQLRecord(rsSQL)
    
    Select Case strCmd
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        '
                
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 And mbytMode = 2 Then blnAllowModify = False
        
        txt(5).Locked = True
        txt(2).Locked = True
        txt(0).Locked = Not blnAllowModify
        txt(7).Locked = Not blnAllowModify
        txt(4).Locked = Not blnAllowModify
        txt(8).Locked = Not blnAllowModify
        txt(9).Locked = True
        txt(1).Locked = True
        txt(6).Locked = True
        txt(3).Locked = True
        
        If txt(0).Locked = False Then
           txt(0).Locked = Not mblnAuditEnter
        End If
        
        If blnAllowModify Then
            If IsPrivs(mstrPrivs, "院级反馈") And IsPrivs(mstrPrivs, "科级反馈") Then
'                cbo(2).ListIndex = 0
                cbo(2).Locked = False
            Else
               If IsPrivs(mstrPrivs, "院级反馈") Then
'                  cbo(2).ListIndex = 0
                  cbo(2).Locked = True
               Else
'                  cbo(2).ListIndex = 1
                  cbo(2).Locked = True
               End If
            End If
        Else
            cbo(2).Locked = Not blnAllowModify
        End If
        
        cbo(3).Locked = Not blnAllowModify
        
        
        cmd(0).Enabled = (blnAllowModify And tbrFree.Buttons("自由").Value = tbrUnpressed)
        tbrFree.Buttons("自由").Enabled = mblnAuditEnter And blnAllowModify
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        txt(0).MaxLength = GetMaxLength("病案反馈记录", "反馈意见")
        txt(1).MaxLength = GetMaxLength("病案反馈记录", "处理说明")
        txt(8).MaxLength = GetMaxLength("病案反馈记录", "补充说明")
        With cbo(0)
            .Clear
            .AddItem "首页记录": .ItemData(.NewIndex) = 5
            .AddItem "住院医嘱": .ItemData(.NewIndex) = 1
            .AddItem "住院病历": .ItemData(.NewIndex) = 2
            .AddItem "护理病历": .ItemData(.NewIndex) = 3
            .AddItem "护理记录": .ItemData(.NewIndex) = 4
            .AddItem "医嘱报告": .ItemData(.NewIndex) = 6
            .AddItem "疾病证明": .ItemData(.NewIndex) = 7
            .AddItem "知情文件": .ItemData(.NewIndex) = 8
            .AddItem "临床路径": .ItemData(.NewIndex) = 9
            mblnReadCom = True
            .ListIndex = 0
            mblnReadCom = False
        End With
        
        With cbo(2)
            .Clear
            .AddItem "院级反馈": .ItemData(.NewIndex) = 0
            .AddItem "科级反馈": .ItemData(.NewIndex) = 1
            .ListIndex = 0
        End With
        
        With cbo(3)
            .Clear
            .AddItem "扣分制": .ItemData(.NewIndex) = 0
            .AddItem "否决制": .ItemData(.NewIndex) = 1
            .ListIndex = 0
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        ExecuteCommand = ExecuteCommand("读取数据", Val(varParam(0)))
        GoTo EndHand
        
    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
        
        cbo(0).Locked = False
        cbo(1).Locked = False
'        cbo(2).Locked = True
        txt(5).Text = ""
        txt(2).Text = ""
        txt(0).Text = ""
        txt(7).Text = ""
        txt(1).Text = ""
        txt(6).Text = ""
        txt(3).Text = ""
        txt(4).Text = ""
        txt(8).Text = ""
        txt(9).Text = ""
        cmd(0).Tag = ""
        
        usrSaveItem.反馈意见 = ""
        
    '------------------------------------------------------------------------------------------------------------------
    Case "缺省数据"
        
        Call zlControl.CboLocate(cbo(0), mstrObject)
                
        Call ExecuteCommand("显示文件")
        
        txt(5).Text = gstrDBUser
        txt(2).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        If Val(GetPara("反馈处理期限", mfrmMain.模块号, "7")) = 0 Then
            txt(7).Text = Format(zlDatabase.Currentdate + 7, "yyyy-MM-dd HH:mm:ss")
        Else
            txt(7).Text = Format(zlDatabase.Currentdate + Val(GetPara("反馈处理期限", mfrmMain.模块号, "7")), "yyyy-MM-dd HH:mm:ss")
        End If
        
        txt(1).Text = ""
        txt(3).Text = ""
        txt(6).Text = ""
        txt(4).Text = ""
        txt(9).Text = mlng次数
        
       
        If IsPrivs(mstrPrivs, "院级反馈") And IsPrivs(mstrPrivs, "科级反馈") Then
            cbo(2).ListIndex = 0
            cbo(2).Locked = False
        Else
           If IsPrivs(mstrPrivs, "院级反馈") Then
              cbo(2).ListIndex = 0
              cbo(2).Locked = True
           Else
              cbo(2).ListIndex = 1
              cbo(2).Locked = True
           End If
        End If
        
        cbo(3).ListIndex = 0
    '------------------------------------------------------------------------------------------------------------------
    Case "显示文件"
    
        With cbo(1)
            .Clear
            .AddItem ""
            If Val(mstr文件ID) > 0 Then
                Select Case mstrObject
                Case "护理记录"
                    Set rs = GetEPRFileStruct(Val(mstr文件ID))
                Case "医嘱报告"
                    Set rs = GetEPRFile(Val(mstr文件ID), mlng医嘱id)
                Case Else
                    If IsNumeric(mstr文件ID) Then
                        Set rs = GetEPRFile(Val(mstr文件ID))
                    Else
                        If Not gobjEmr Is Nothing Then
                            Set rs = GetEMRFile(Val(mstr文件ID))
                        End If
                    End If
                End Select
                
                If rs.BOF = False Then
                    Do While Not rs.EOF
                        .AddItem rs("名称").Value
                        If IsNumeric(mstr文件ID) Then
                            .ItemData(.NewIndex) = mstr文件ID
                        Else
                            .Tag = mstr文件ID
                        End If
                        rs.MoveNext
                    Loop
                    .ListIndex = 1
                End If
                
            Else
                .ListIndex = 0
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
        
        Call ExecuteCommand("清空数据")
        mblnReading = True
        
        Set rs = GetQuestion(mrsCondition, "", 1, Val(varParam(0)))
        If rs.BOF = False Then
        
            txt(0).Text = zlCommFun.NVL(rs("反馈意见").Value)
            cmd(0).Tag = zlCommFun.NVL(rs("反馈项目ID").Value)
            
            If Val(cmd(0).Tag) > 0 Then usrSaveItem.反馈意见 = zlCommFun.NVL(rs("反馈意见").Value)
            
            txt(1).Text = zlCommFun.NVL(rs("处理说明").Value)
            txt(2).Text = Format(zlCommFun.NVL(rs("反馈时间").Value), "yyyy-MM-dd HH:mm:ss")
            txt(3).Text = Format(zlCommFun.NVL(rs("处理时间").Value), "yyyy-MM-dd HH:mm:ss")
            txt(7).Text = Format(zlCommFun.NVL(rs("处理期限").Value), "yyyy-MM-dd HH:mm:ss")
            txt(6).Text = zlCommFun.NVL(rs("处理人").Value)
            txt(5).Text = zlCommFun.NVL(rs("反馈人").Value)
            txt(4).Text = zlCommFun.NVL(rs("分值").Value)
            txt(8).Text = zlCommFun.NVL(rs("补充说明").Value)
            txt(9).Text = zlCommFun.NVL(rs("反馈次数").Value)
            
            cbo(2).ListIndex = zlCommFun.NVL(rs("评分级别").Value, 0)
            cbo(3).ListIndex = zlCommFun.NVL(rs("分制").Value, 0)
            
            mstr文件ID = zlCommFun.NVL(rs("文件id").Value, "0")
            mlng医嘱id = zlCommFun.NVL(rs("医嘱id").Value, 0)
            mlng科室ID = zlCommFun.NVL(rs("科室id").Value, 0)
            
            mstrObject = zlCommFun.NVL(rs("反馈对象").Value)
            
            Call zlControl.CboLocate(cbo(0), zlCommFun.NVL(rs("反馈对象").Value))
            
            cbo(1).Text = zlCommFun.NVL(rs("反馈记录").Value)
            If mstr文件ID <> "0" Then
                Call ExecuteCommand("显示文件")
            End If
        End If
        
    End Select

    ExecuteCommand = True
    
    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:
    mblnReading = False
End Function

Private Sub cbo_Click(Index As Integer)
    Dim lngRow As Long
    Select Case Index
    Case 0
        Select Case cbo(Index).ItemData(cbo(Index).ListIndex)
        Case 1, 5           '医嘱,首页
            cbo(1).Clear
            If mblnReadCom Then Exit Sub
            RaiseEvent AfterParaments(cbo(0).Text, GetTypeAuditObject(cbo(1).Text))
        Case Else
            cbo(1).Clear
            cbo(1).AddItem ""
            Erase mTypeAuditObject()
            
            mRsType.Filter = "上级id='R" & cbo(Index).ItemData(cbo(Index).ListIndex) & "'"
            If mRsType.BOF = False Then
                 mRsType.MoveFirst
                 Do While Not mRsType.EOF
                    ReDim Preserve mTypeAuditObject(lngRow)
                    cbo(1).AddItem mRsType("名称").Value, cbo(1).ItemData(cbo(1).NewIndex) = zlCommFun.NVL(mRsType("参数").Value)
                    mTypeAuditObject(lngRow).strName = mRsType("名称").Value
                    mTypeAuditObject(lngRow).strID = mRsType("ID").Value
                    mTypeAuditObject(lngRow).strPara = mRsType("参数").Value
                    lngRow = lngRow + 1
                    mRsType.MoveNext
                 Loop
            End If
            
            On Error Resume Next
            mrsEmr.Filter = "上级id='R" & cbo(Index).ItemData(cbo(Index).ListIndex) & "'"
            If ObjPtr(mrsEmr) > 0 Then
            If mrsEmr.BOF = False Then
                 mrsEmr.MoveFirst
                 Do While Not mrsEmr.EOF
                    ReDim Preserve mTypeAuditObject(lngRow)
                    cbo(1).AddItem mrsEmr("名称").Value, cbo(1).ItemData(cbo(1).NewIndex) = zlCommFun.NVL(mrsEmr("参数").Value)
                    mTypeAuditObject(lngRow).strName = mrsEmr("名称").Value
                    mTypeAuditObject(lngRow).strID = mrsEmr("ID").Value
                    mTypeAuditObject(lngRow).strPara = mrsEmr("参数").Value
                    lngRow = lngRow + 1
                    mrsEmr.MoveNext
                 Loop
            End If
            End If
            
            cbo(1).ListIndex = 0
        End Select
    Case 1
        RaiseEvent AfterParaments(cbo(0).Text, GetTypeAuditObject(cbo(1).Text))
    Case 2
         If cbo(Index).Text = "院级反馈" Then
            RaiseEvent AfterQuestionType(True)
         Else
            RaiseEvent AfterQuestionType(False)
         End If
         DataChanged = True
    End Select
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim intObject       As Integer
    Dim strFileID       As String
    
    On Error GoTo ErrH
    '适用对象
    intObject = cbo(0).ItemData(cbo(0).ListIndex)
    '文件ID
    If cbo(1).ListIndex >= 0 Then
        strFileID = cbo(1).ItemData(cbo(1).ListIndex)
        If strFileID = 0 Then
            strFileID = cbo(1).Tag
            If strFileID = "" Then
                strFileID = 0
            End If
        End If
    End If
    
    Select Case Index
        '------------------------------------------------------------------------------------------------------------------
        Case 0      '反馈意见(来源于病案评分标准)
            Call GetAuditItem(intObject, strFileID)
    End Select
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
'    picPane(1).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
'    fra.BackColor = COLOR_NativeXpPlain.SpecialGroupClient
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picPane(1).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set mrsCondition = Nothing
    Set mRsType = Nothing
    Set zlCheck = Nothing
    Erase mTypeAuditObject
    Set mrsEmr = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
        Case 1
            fra.Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
            cbo(1).Move cbo(1).Left, cbo(1).Top, fra.Width - cbo(1).Left - 60
            txt(0).Move txt(0).Left, txt(0).Top, fra.Width - txt(0).Left - 60 - cmd(0).Width - 15
            cmd(0).Move txt(0).Left + txt(0).Width + 15, cmd(0).Top
            txt(5).Move txt(5).Left, txt(5).Top, fra.Width - txt(5).Left - 60 - cmd(0).Width - 15
            txt(1).Move txt(1).Left, txt(1).Top, fra.Width - txt(1).Left - 60 - cmd(0).Width - 15
            txt(2).Move txt(2).Left, txt(2).Top, fra.Width - txt(2).Left - 60 - cmd(0).Width - 15
            txt(3).Move txt(3).Left, txt(3).Top, fra.Width - txt(3).Left - 60 - cmd(0).Width - 15
            txt(6).Move txt(6).Left, txt(6).Top, fra.Width - txt(6).Left - 60 - cmd(0).Width - 15
            txt(7).Move txt(7).Left, txt(7).Top, fra.Width - txt(7).Left - 60 - cmd(0).Width - 15
            txt(8).Move txt(8).Left, txt(8).Top, fra.Width - txt(8).Left - 60 - cmd(0).Width - 15
            
            txt(4).Move txt(4).Left, txt(4).Top, fra.Width - txt(4).Left - 60 - cmd(0).Width - 15
            txt(9).Move txt(9).Left, txt(9).Top, fra.Width - txt(9).Left - 60 - cmd(0).Width - 15
            
            cbo(2).Move cbo(3).Left, txt(9).Top  ', fra.Width - cbo(2).Left - 60 - cmd(0).Width - 15
            
            
            
    End Select
End Sub

Private Sub tbrFree_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call ExecuteCommand("控件状态")
End Sub

Private Sub txt_Change(Index As Integer)
        
    If mblnReading Then Exit Sub
    
    DataChanged = True
 
    If (Index = 0 Or Index = 4 Or Index = 8) And cmd(0).Enabled Then
        txt(Index).Tag = "Changed"
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 0, 1, 4, 8
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim StrText As String, strFileID As String
    
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
            If cmd(0).Enabled Then
                txt(Index).Tag = ""
                StrText = UCase(txt(Index).Text)
                If cbo(1).ListIndex = -1 Then
                    strFileID = ""
                Else
                    strFileID = CStr(cbo(1).ItemData(cbo(1).ListIndex))
                End If
                If strFileID = "0" Then
                    strFileID = cbo(1).Tag
                    If strFileID = "" Then
                        strFileID = "0"
                    End If
                End If
                Call GetAuditItem(cbo(0).ItemData(cbo(0).ListIndex), strFileID, StrText)
            Else
                txt(Index).Tag = ""
                If Index = 3 Then
                    Call LocationObj(cbo(0))
                Else
                    zlCommFun.PressKey vbKeyTab
                End If
            End If

        '--------------------------------------------------------------------------------------------------------------
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If Index = 4 Then
            If KeyAscii = 46 And InStrRev(txt(Index).Text, ".", -1) > 0 Then KeyAscii = 0
            If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then KeyAscii = 0
        End If
    End If

End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 0, 1, 8
        zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub

    Select Case Index
        Case 0
            If (txt(Index).Tag = "Changed") And cmd(0).Enabled Then
                txt(Index).Text = IIf(usrSaveItem.反馈意见 = "", txt(Index).Text, usrSaveItem.反馈意见)
                txt(Index).Tag = ""
            End If
        Case 7
            If (txt(Index).Tag = "Changed") And cmd(0).Enabled Then
                txt(Index).Text = IIf(usrSaveItem.反馈意见 = "", txt(Index).Text, usrSaveItem.反馈意见)
                txt(Index).Tag = ""
            End If
    End Select
End Sub

Private Sub GetAuditItem(intObject As Integer, strFileID As String, Optional shortName As String = "")
Dim rsData As ADODB.Recordset, strSubid As String, strReturn As String
On Error GoTo ErrH
    If IsNumeric(strFileID) Then
        '检测文件ID存在与否
        '电子病案记录 如果存在则直接取文件ID关联，否则直接按类型读取
        gstrSQL = "Select 文件ID From 电子病历记录 a Where a.ID = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(strFileID))
        If zlCheck.Connection_ChkRsState(rsData) Then
            strFileID = 0
        Else
            strFileID = "" & rsData.Fields!文件ID
        End If
    Else
        If Not gobjEmr Is Nothing Then
            If InStr(strFileID, "|") > 0 Then
                strSubid = Split(strFileID, "|")(1)
                strFileID = Split(strFileID, "|")(0)
            End If
            gstrSQL = "Select RawtoHex(Antetype_Id) 文件ID From Bz_Doc_Tasks A Where Real_Doc_Id = Hextoraw(:docid)" & IIf(strSubid = "", "", " And Subdoc_Id =:subdocid")
            strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strFileID & "^16^docid" & IIf(strSubid = "", "", "|" & strSubid & "^16^subdocid"), rsData)
            If strReturn <> "" Then strFileID = 0
            If Not rsData Is Nothing Then
            If rsData.RecordCount > 0 Then
                strFileID = rsData!文件ID
            End If
            End If
        End If
    End If
        
    If strFileID = "0" Then
        gstrSQL = "Select A.ID,A.名称,A.说明, A.分制,A.分值 From 病案审查目录  A ,病案审查分类 B,病案审查方案 C where  A.分类ID =B.id And B.方案ID =C.ID And C.启用时间 is Not Null And A.适用对象 = " & CStr(intObject)
    Else
        gstrSQL = "Select A.ID,A.名称,A.说明,A.分制,A.分值 From 病案审查目录 A ,病案审查分类 B,病案审查方案 C  where A.分类ID =B.id And B.方案ID =C.ID And C.启用时间 is Not Null And A.适用对象 = " & CStr(intObject) & " And (A.文件ID is null or instr(','|| A.文件ID || ',' , ','|| '" & strFileID & "' || ',')>0 )"
    '20100906 --zq nvl(文件ID,'')=''
    End If
    If shortName <> "" Then
        gstrSQL = gstrSQL & vbCrLf & "And (A.编码 like '%" & shortName & "%' or A.简码 like '%" & shortName & "%' or A.名称 like '%" & shortName & "%')"
    End If
    
    Set rsData = zlDatabase.ShowSelect(Me, gstrSQL, 0)
    If zlCheck.Connection_ChkRsState(rsData) Then
        RaiseEvent RefStatus
        Exit Sub
    Else
        RaiseEvent RefStatus
    End If
    
    
    txt(0).Text = zlCommFun.NVL(rsData("名称").Value)
    cmd(0).Tag = zlCommFun.NVL(rsData("ID").Value)
    txt(0).Tag = ""
    usrSaveItem.反馈意见 = txt(0).Text
    
    If zlCommFun.NVL(rsData("分制").Value, 0) = 0 Then
        '扣分制
        cbo(3).ListIndex = 0
        txt(4).Text = zlCommFun.NVL(rsData("分值").Value, 0)
        mlng分值 = zlCommFun.NVL(rsData("分值").Value, 0)
    Else
        '否决制
        cbo(3).ListIndex = 1
        txt(4).Text = 0
        mlng分值 = 0
    End If
    
    
    DataChanged = True
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
End Sub

Public Function Init病案类型(ByVal rs As ADODB.Recordset, ByVal rsEmr As ADODB.Recordset)
    Set mRsType = rs
    Set mrsEmr = rsEmr
    cbo(0).ListIndex = 0
    cbo(1).Clear
End Function

Private Function GetTypeAuditObject(ByVal strName As String) As String
    '根据对象名称从数组中获取参数值
    Dim i As Integer
    If strName = "" Then Exit Function
    
    For i = 0 To UBound(mTypeAuditObject)
        If mTypeAuditObject(i).strName = strName Then
            GetTypeAuditObject = mTypeAuditObject(i).strPara
            Exit Function
        End If
    Next
    
End Function

Public Function RefActiveFrom()
    Me.AutoRedraw = True
    Me.Enabled = True
        
End Function
