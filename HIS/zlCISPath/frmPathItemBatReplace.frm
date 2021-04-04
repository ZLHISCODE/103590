VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathItemBatReplace 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "项目批量调整"
   ClientHeight    =   9855
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13725
   FillColor       =   &H00404040&
   ForeColor       =   &H8000000C&
   Icon            =   "frmPathItemBatReplace.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10521.35
   ScaleMode       =   0  'User
   ScaleWidth      =   13973.42
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSplit 
      BorderStyle     =   0  'None
      Height          =   42
      Left            =   4200
      TabIndex        =   40
      Top             =   5280
      Width           =   4695
   End
   Begin VB.PictureBox picAdvice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   4680
      ScaleHeight     =   2145
      ScaleWidth      =   6615
      TabIndex        =   37
      Top             =   5520
      Width           =   6615
      Begin VB.CommandButton cmdEdit 
         Caption         =   "替换项目编辑"
         Height          =   420
         Left            =   0
         TabIndex        =   39
         Top             =   120
         Width           =   1500
      End
      Begin zlCISPath.UCAdviceList ucAdvice 
         Height          =   1575
         Left            =   0
         TabIndex        =   38
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2778
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   13725
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9180
      Width           =   13725
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H8000000E&
         Caption         =   "退出(&Q)"
         Height          =   420
         Left            =   12000
         TabIndex        =   9
         Top             =   120
         Width           =   1500
      End
      Begin VB.CommandButton cmdBatExe 
         BackColor       =   &H80000014&
         Caption         =   "批量替换(&B)"
         Height          =   420
         Left            =   10440
         TabIndex        =   8
         Top             =   120
         Width           =   1500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   20400
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   20280
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4680
      ScaleHeight     =   975
      ScaleWidth      =   9735
      TabIndex        =   19
      Top             =   960
      Width           =   9735
      Begin VB.PictureBox picItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   6855
         TabIndex        =   20
         Top             =   0
         Width           =   6855
         Begin VB.CommandButton cmd用法 
            Height          =   240
            Left            =   4440
            Picture         =   "frmPathItemBatReplace.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   165
            Width           =   270
         End
         Begin VB.CommandButton cmd频率 
            Height          =   240
            Left            =   4440
            Picture         =   "frmPathItemBatReplace.frx":6948
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   570
            Width           =   270
         End
         Begin VB.TextBox txt单量 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   525
            MaxLength       =   10
            TabIndex        =   0
            Top             =   135
            Width           =   1290
         End
         Begin VB.TextBox txt总量 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   525
            MaxLength       =   10
            TabIndex        =   2
            Top             =   540
            Width           =   1290
         End
         Begin VB.CheckBox chkPra 
            BackColor       =   &H80000005&
            Caption         =   "只替换频率相同的"
            Height          =   255
            Index           =   2
            Left            =   5040
            TabIndex        =   6
            Top             =   600
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkPra 
            BackColor       =   &H80000005&
            Caption         =   "只替换用法相同的"
            Height          =   255
            Index           =   1
            Left            =   5040
            TabIndex        =   5
            Top             =   360
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkPra 
            BackColor       =   &H80000005&
            Caption         =   "只替换用量相同的"
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   4
            Top             =   120
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox txt频率 
            Height          =   300
            Left            =   2925
            TabIndex        =   3
            Top             =   540
            Width           =   1815
         End
         Begin VB.TextBox txt用法 
            Height          =   300
            Left            =   2925
            TabIndex        =   1
            Top             =   135
            Width           =   1815
         End
         Begin VB.Label lbl单量单位 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "g"
            Height          =   180
            Left            =   1935
            TabIndex        =   26
            Top             =   195
            Width           =   405
         End
         Begin VB.Label lbl单量 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单量"
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lbl总量 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "总量"
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl总量单位 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "盒"
            Height          =   180
            Left            =   1920
            TabIndex        =   23
            Top             =   600
            Width           =   450
         End
         Begin VB.Label lbl用法 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "用法"
            Height          =   180
            Left            =   2520
            TabIndex        =   22
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lbl频率 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "频率"
            Height          =   180
            Left            =   2520
            TabIndex        =   21
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H80000005&
         Caption         =   "查找路径(&F)"
         Height          =   420
         Left            =   6960
         TabIndex        =   7
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.PictureBox picSplit 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   3840
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7335
      ScaleWidth      =   45
      TabIndex        =   18
      Top             =   1200
      Width           =   45
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   13725
      TabIndex        =   16
      Top             =   0
      Width           =   13725
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "6、点击批量替换后完成替换。"
         Height          =   255
         Index           =   8
         Left            =   8040
         TabIndex        =   36
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "3、查找包含选择的诊疗项目的路径。"
         Height          =   255
         Index           =   7
         Left            =   8040
         TabIndex        =   35
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathItemBatReplace.frx":6A3E
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "5、设置诊疗项目对应的替换项目。"
         Height          =   255
         Index           =   6
         Left            =   4560
         TabIndex        =   31
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "说明:"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   30
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "2、设置诊疗项目替换的规则。"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   28
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "4、勾选需要替换的路径表单。"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   27
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "1、选择需要替换的诊疗项目。"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   17
         Top             =   120
         Width           =   3375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   20280
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   4680
      ScaleHeight     =   2775
      ScaleWidth      =   7335
      TabIndex        =   12
      Top             =   2280
      Width           =   7335
      Begin XtremeReportControl.ReportControl rptPath 
         Height          =   2055
         Left            =   0
         TabIndex        =   13
         Top             =   240
         Width           =   7215
         _Version        =   589884
         _ExtentX        =   12726
         _ExtentY        =   3625
         _StockProps     =   0
         BorderStyle     =   2
      End
      Begin VB.Label lblNote 
         BackColor       =   &H80000005&
         Caption         =   "路径列表"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6135
      ScaleWidth      =   4575
      TabIndex        =   10
      Top             =   840
      Width           =   4575
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4215
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   3255
         _Version        =   589884
         _ExtentX        =   5741
         _ExtentY        =   7435
         _StockProps     =   0
         BorderStyle     =   2
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.Frame fraFind 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   900
         Left            =   120
         TabIndex        =   41
         Top             =   135
         Width           =   4455
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   480
            TabIndex        =   44
            ToolTipText     =   "查找下一个(F3)"
            Top             =   480
            Width           =   3255
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "路径查找"
            Height          =   300
            Index           =   1
            Left            =   2760
            TabIndex        =   43
            Top             =   60
            Width           =   1215
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "直接查找"
            Height          =   300
            Index           =   0
            Left            =   0
            TabIndex        =   42
            Top             =   60
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblFind 
            BackColor       =   &H80000005&
            Caption         =   "查找"
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   510
            Width           =   495
         End
      End
      Begin VB.Label lblStopNote 
         BackColor       =   &H80000005&
         Caption         =   "路径表中用到的已停用项目"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   2295
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2880
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathItemBatReplace.frx":7193
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathItemBatReplace.frx":772D
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "1、选中停用项目后，设置过滤条件后查找满足条件的路径"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmPathItemBatReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrSelItem As String    '记录选中项目 格式:诊疗项目ID_收费细目ID
Private mstrPrivs As String      '临床路径模块权限
Private mrsAdvice As ADODB.Recordset   '替换项目医嘱记录集
Private mblnChange As Boolean   '标识过滤参数值加载后的有无变动情况 T-变动，F-未变动过

'---------------------------------
Private Enum CHK_INDEX
    chk_用量 = 0
    chk_用法 = 1
    chk_频率 = 2
End Enum

Private Enum COL_LIST
    COL_分类 = 0
    COL_编码
    COL_名称
    COL_商品名
    COL_产地
    COL_药品剂型
    
    '隐藏列
    COL_诊疗项目ID
    COL_收费细目ID
    COL_诊疗类别
    COL_操作类型
    COL_执行分类
    COL_计算方式
    COL_标本部位
    COL_检查方法
    COL_内容ID
    COL_相关ID
    COL_简码
End Enum

Private Enum COL_PATH
    Path_ID = 0
    Path_选择
    Path_分类
    Path_编码
    Path_名称
    Path_版本
    Path_说明
End Enum

Private Enum CONST_COLOR
    Color_Enabled = &H80000005
    Color_UNEnabled = &H8000000F
End Enum
'-------------------------------------------------------------------------------------------------------
Public Sub ShowMe(frmParent As Object, ByVal strPrivs As String)
'功能:入口函数
'参数:主窗体
'
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
End Sub

Private Sub chkPra_Click(Index As Integer)
    Dim blnCheck As Boolean
    
    blnCheck = chkPra(Index) = vbChecked
    If Index = chk_用量 Then
        SetEditable IIf(blnCheck, 1, -1), IIf(blnCheck, 1, -1)
    ElseIf Index = chk_用法 Then
        SetEditable , , IIf(blnCheck, 1, -1)
    ElseIf Index = chk_频率 Then
        SetEditable , , , IIf(blnCheck, 1, -1)
    End If
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub cmdBatExe_Click()
'功能:批量替换处理
    '保存数据前的检查控制
    Dim i As Long
    Dim strPath As String
    Dim strTmp As String
    
    If mrsAdvice.RecordCount = 0 Then
        MsgBox "请先设置替换项目,再执行【批量替换】功能。", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    With rptPath
        For i = 1 To .Rows.count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).Record(Path_选择).Checked Then
                    strPath = strPath & ":" & .Rows(i).Record(Path_ID).Value & "," & .Rows(i).Record(Path_版本).Value
                    If .Rows(i).Record(Path_版本).Value = .Rows(i).Record.Tag Then  '记录下最新版本的路径ID和版本号
                        strTmp = strTmp & "," & .Rows(i).Record(Path_ID).Value & "_" & .Rows(i).Record.Tag
                    End If
                End If
            End If
        Next
        If strPath <> "" Then strPath = Mid(strPath, 2)
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    End With
    
    If strPath = "" Then
        MsgBox "请先选择替换的路径,再执行【批量替换】功能。", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
        Exit Sub
    End If
    '数据保存
    If SaveData(strPath, strTmp) Then
        '刷新界面
        Call RefreshData
    End If
End Sub

Private Sub cmdEdit_Click()
'功能:替换项目编辑
    Dim rsScheme As ADODB.Recordset
    Dim colAdviceID As New Collection
    Dim lng序号 As Long
    Dim lng医嘱ID As Long
    
    Call InitSchemeRecordset(rsScheme)
    
    Do While Not mrsAdvice.EOF
        rsScheme.AddNew
        rsScheme!序号 = mrsAdvice!ID
        rsScheme!相关序号 = mrsAdvice!相关id
        rsScheme!期效 = mrsAdvice!期效
        rsScheme!诊疗项目ID = mrsAdvice!诊疗项目ID
        rsScheme!收费细目ID = mrsAdvice!收费细目ID
        rsScheme!医嘱内容 = mrsAdvice!医嘱内容
        rsScheme!单次用量 = mrsAdvice!单次用量
        rsScheme!总给予量 = mrsAdvice!总给予量
        rsScheme!医生嘱托 = mrsAdvice!医生嘱托
        rsScheme!执行频次 = mrsAdvice!执行频次
        rsScheme!频率次数 = mrsAdvice!频率次数
        rsScheme!频率间隔 = mrsAdvice!频率间隔
        rsScheme!间隔单位 = mrsAdvice!间隔单位
        rsScheme!时间方案 = mrsAdvice!时间方案
        rsScheme!执行科室ID = mrsAdvice!执行科室ID
        rsScheme!执行性质 = mrsAdvice!执行性质
        rsScheme!标本部位 = mrsAdvice!标本部位
        rsScheme!检查方法 = mrsAdvice!检查方法
        rsScheme!配方ID = mrsAdvice!配方ID
        rsScheme!组合项目ID = mrsAdvice!组合项目ID
        rsScheme!执行标记 = mrsAdvice!执行标记
        
        rsScheme.Update
        mrsAdvice.MoveNext
    Loop
    
    Set rsScheme = gobjKernel.ShowSchemeEdit(Me, 2, rsScheme, False, False, "", 2, rptList.SelectedRows(0).Record(COL_诊疗类别).Value & "", _
                    rptList.SelectedRows(0).Record(COL_操作类型).Value & "", rptList.SelectedRows(0).Record(COL_执行分类).Value & "")
    
    
    '先删除以前的医嘱ID
    If mrsAdvice.RecordCount > 0 And Not rsScheme Is Nothing Then
        Call InitAdviceRecordset '重新初始化
    End If

    If Not rsScheme Is Nothing Then
         '先产生新的医嘱ID
        Do While Not rsScheme.EOF
            lng医嘱ID = zlDatabase.GetNextId("路径医嘱内容")
            colAdviceID.Add lng医嘱ID, "_" & rsScheme!序号
            rsScheme.MoveNext
        Loop
        rsScheme.MoveFirst: lng序号 = 1
        Do While Not rsScheme.EOF
            mrsAdvice.AddNew
            mrsAdvice!ID = colAdviceID("_" & rsScheme!序号)
            If Not IsNull(rsScheme!相关序号) Then
                mrsAdvice!相关id = colAdviceID("_" & rsScheme!相关序号)
            End If
            mrsAdvice!序号 = lng序号
            mrsAdvice!期效 = rsScheme!期效
            mrsAdvice!诊疗项目ID = rsScheme!诊疗项目ID
            mrsAdvice!收费细目ID = rsScheme!收费细目ID
            If IsNull(rsScheme!诊疗项目ID) Then
                mrsAdvice!医嘱内容 = rsScheme!医嘱内容 '自由录入医嘱才保存
            End If
            mrsAdvice!单次用量 = rsScheme!单次用量
            mrsAdvice!总给予量 = rsScheme!总给予量
            mrsAdvice!医生嘱托 = rsScheme!医生嘱托
            mrsAdvice!执行频次 = rsScheme!执行频次
            mrsAdvice!频率次数 = rsScheme!频率次数
            mrsAdvice!频率间隔 = rsScheme!频率间隔
            mrsAdvice!间隔单位 = rsScheme!间隔单位
            mrsAdvice!时间方案 = rsScheme!时间方案
            mrsAdvice!执行科室ID = rsScheme!执行科室ID
            mrsAdvice!执行性质 = rsScheme!执行性质
            mrsAdvice!标本部位 = rsScheme!标本部位
            mrsAdvice!检查方法 = rsScheme!检查方法
            mrsAdvice!是否缺省 = rsScheme!是否缺省
            mrsAdvice!是否备选 = rsScheme!是否备选
            mrsAdvice!配方ID = rsScheme!配方ID
            mrsAdvice!组合项目ID = rsScheme!组合项目ID
            mrsAdvice!执行标记 = rsScheme!执行标记
            
            mrsAdvice.Update
            
            lng序号 = lng序号 + 1
            rsScheme.MoveNext
        Loop
        If mrsAdvice.RecordCount > 1 Then mrsAdvice.MoveFirst
    End If
    
    Call ShowAdvice
    cmdBatExe.Enabled = True
End Sub

Private Sub cmdFind_Click()
    Dim strSql As String
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objCol As ReportColumn
    Dim str类别 As String
    
    Dim i As Long
    '清空数据
    If rptList.Records.count = 0 Then Exit Sub
    If rptList.SelectedRows.count < 1 Then Exit Sub
    
    rptPath.Records.DeleteAll: rptPath.Populate: cmdEdit.Enabled = False
    
    On Error GoTo errH
    strSql = "Select Distinct d.Id, d.分类, d.编码, d.名称,d.说明,H.版本号,D.最新版本 " & vbNewLine & _
            "From 路径医嘱内容 A, 临床路径医嘱 B, 临床路径项目 C, 临床路径版本 H,临床路径目录 D" & vbNewLine & _
            "Where a.Id = b.医嘱内容id And b.路径项目id = c.Id And c.路径id = H.路径Id And c.版本号 = H.版本号 And H.停用人 is null And H.路径Id=D.ID"
    With rptList.SelectedRows(0)
    
        If Val(.Record(COL_收费细目ID).Value) = 0 Then
            strSql = strSql & " And a.诊疗项目ID =[1]"
        Else
            strSql = strSql & " And a.收费细目ID =[1]"
        End If
        str类别 = .Record(COL_诊疗类别).Value
        If InStr(",D,C,", "," & str类别 & ",") > 0 Then
            strSql = strSql & " And Instr([7], ',' || NVl(a.相关ID,a.Id)|| ',') > 0 "
        End If
        
        If chkPra(chk_用量).Value = vbChecked Then
            If txt单量.Text <> "" Then
                strSql = strSql & " And a.单次用量 =[2] "
            End If
            If txt总量.Text <> "" Then
                strSql = strSql & " And a.总给予量 = [3] "
            End If
        End If
        
        If chkPra(chk_用法).Value = vbChecked Then
            If txt用法.Text <> "" Then
                strSql = strSql & " and exists (select 1 from 路径医嘱内容  E where e.id=a.相关id and e.诊疗项目id = [4]) "
            End If
        End If
        
        If chkPra(chk_频率).Value = vbChecked Then
            If txt频率.Text <> "" Then
                strSql = strSql & " And a.执行频次 =[5] "
            End If
        End If
            
        If InStr(mstrPrivs, "全院路径") = 0 Then
            '没有权限时，只能对只应用于本科的路径进行处理
            strSql = strSql & _
                     " And D.通用 = 2 And Exists" & vbNewLine & _
                     "      (Select 1 From 部门人员 E,临床路径科室 F  " & vbNewLine & _
                     "       Where E.人员id = [6] And F.科室id = E.部门id And F.路径id = D.ID)"
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(Val(.Record(COL_收费细目ID).Value) = 0, Val(.Record(COL_诊疗项目ID).Value), Val(.Record(COL_收费细目ID).Value)), _
                    Val(txt单量.Text), Val(txt总量.Text), Val(txt用法.Tag), txt频率.Text, UserInfo.ID, "," & .Record(COL_内容ID).Value & ",")
        If rsTmp.RecordCount = 0 Then Exit Sub
        
        With rptPath
            For i = 1 To rsTmp.RecordCount
                Set objRecord = .Records.Add
                objRecord.AddItem rsTmp!ID & ""
                
                Set objItem = objRecord.AddItem("")
                objItem.HasCheckbox = True
                If .Columns(Path_选择).Icon = img16.ListImages("UnCheck").Index - 1 Then
                    objItem.Checked = True
                Else
                    objItem.Checked = False
                End If
                objRecord.AddItem rsTmp!分类 & ""
                objRecord.AddItem rsTmp!编码 & ""
                objRecord.AddItem rsTmp!名称 & ""
                objRecord.AddItem rsTmp!版本号 & ""
                objRecord.AddItem rsTmp!说明 & ""
                objRecord.Tag = rsTmp!最新版本 & ""  '依据该值判断变动记录是否插入
                rsTmp.MoveNext
            Next
            .Populate
        End With
        
        cmdEdit.Enabled = rptPath.Records.count > 0
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmd频率_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str范围 As String, int频率 As Integer, vRect As RECT
    Dim lng诊疗项目ID As Long
       
    If rptList.SelectedRows.count = 0 Then Exit Sub  '非正常情况
    With rptList.SelectedRows(0)
         strSql = "Select 执行频率 From 诊疗项目目录 Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_诊疗项目ID).Value))
        If Not rsTmp.EOF Then int频率 = NVL(rsTmp!执行频率, 0)
        
        If txt总量.Text <> "" Then '临嘱
            If .Record(COL_诊疗类别).Value <> "7" And int频率 = 0 Then
                str范围 = "1,-1" '临嘱可以为一次性
            Else
                str范围 = Get频率范围(int频率)
            End If
        Else
            str范围 = Get频率范围(int频率)
            int频率 = Decode(str范围, "1", 0, "2", 0, "-1", 1, "-2", 2, "-3", 1, "-5", 1)
        End If
        
        '可选择频率的常用频率
        lng诊疗项目ID = Val(.Record(COL_诊疗项目ID).Value)
        strSql = ""
        If InStr("," & str范围 & ",", ",1,") > 0 Then
            strSql = " And (Exists(Select 1 From 诊疗用法用量 Where 项目ID=[2] And 用法ID is NULL And 频次=A.编码 And A.适用范围=1)" & _
                " Or (Select Count(*) From 诊疗用法用量 Where 项目ID=[2] And 用法ID is NULL And 频次 Is Not NULL)<=1)"
        End If
        strSql = _
            " Select Rownum as ID,A.编码,A.名称,A.简码," & _
            " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位,A.适用范围 as 范围ID" & _
            " From 诊疗频率项目 A" & _
            " Where (Instr([1],','||A.适用范围||',')>0  Or a.适用范围=[3])" & strSql & _
            " Order by A.适用范围,A.编码"
        vRect = zlControl.GetControlRect(txt频率.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗频率", False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt频率.Height, blnCancel, False, True, "," & str范围 & ",", lng诊疗项目ID, IIf(txt总量.Text <> "", -5, -3))
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有可用的诊疗频率项目，请先到医嘱频率管理中设置。", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txt频率)
            txt频率.SetFocus: Exit Sub
        End If
        txt频率.Text = rsTmp!名称 & ""
        Call zlControl.TxtSelAll(txt频率)
        txt频率.SetFocus
  
    End With
End Sub

Private Sub cmd用法_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int类型 As Integer, vRect As RECT
    
    If rptList.SelectedRows.count = 0 Then Exit Sub  '非正常情况
    
    With rptList.SelectedRows(0)
        If InStr(",5,6,", .Record(COL_诊疗类别).Value) > 0 Then
            int类型 = 2 '给药途径
        ElseIf .Record(COL_诊疗类别).Value = "C" Then
            int类型 = 6 '采集方法
        ElseIf .Record(COL_诊疗类别).Value = "K" Then
            int类型 = 8 '输血途径
        Else
            int类型 = 4 '中药用法
        End If
        If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
            strSql = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[2] And 性质>0)" & _
                " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                    " Where A.用法ID=B.ID And B.服务对象 IN([3],3) And A.项目ID=[2] And A.性质>0)<=1)"
        End If
        strSql = "Select Distinct A.ID,A.编码,A.名称,C.名称 as 分类" & _
            " From 诊疗项目别名 B,诊疗项目目录 A,诊疗分类目录 C" & _
            " Where A.ID=B.诊疗项目ID And A.分类ID=C.ID(+)" & _
            " And A.类别='E' And A.操作类型=[1] And A.服务对象 IN([3],3)" & strSql & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " Order by A.编码"
        vRect = zlControl.GetControlRect(txt用法.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, lbl用法.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, CStr(int类型), .Record(COL_诊疗项目ID).Value, 2)
        If rsTmp Is Nothing Then
            txt用法.SetFocus: Exit Sub
        End If

        txt用法.SetFocus
        txt用法.Text = rsTmp!名称 & ""
        txt用法.Tag = rsTmp!ID & ""
        Call zlControl.TxtSelAll(txt用法)
    End With
End Sub

Private Sub Form_Load()
    Call InitRPTListColumn
    Call InitRPTPathColumn
    '加载停用项目
    Call RefreshData
    cmdEdit.Enabled = False
    '替换项目初始
    Call InitAdviceTable
    optType(0).Value = True
    Call optType_Click(0)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '分隔线
    With picSplit
        .Left = Me.ScaleWidth \ 3
        .Top = picInfo.Height
        .Width = 45
        .Height = Me.ScaleHeight - picInfo.Height
        
    End With
    '左边栏
    With picLeft
        .Left = 0
        .Top = picInfo.Height
        .Width = picSplit.Left
        .Height = Me.ScaleHeight - picInfo.Height - picBottom.Height
    End With
    '顶部
    With picTop
        .Left = picSplit.Left + picSplit.Width
        .Top = picInfo.Height
        .Width = Me.ScaleWidth - .Left - 45
    End With
    
    '医嘱显示区 高度固定
    With picAdvice
        .Left = picSplit.Left + picSplit.Width
        .Top = Me.ScaleHeight - picBottom.Height - .Height + 30
        .Width = Me.ScaleWidth - .Left - 45 '
    End With
    
    fraSplit.Move picSplit.Left + picSplit.Width, picAdvice.Top + 45, picAdvice.Width, 45
    '中间
    With picMain
        .Left = picSplit.Left + picSplit.Width
        .Top = picInfo.Height + picTop.Height
        .Width = Me.ScaleWidth - .Left - 45
        .Height = Me.ScaleHeight - picInfo.Height - picTop.Height - picBottom.Height - picAdvice.Height - fraSplit.Height
    End With
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'功能:卸载处理

    mstrSelItem = ""
    If Not mrsAdvice Is Nothing Then
        Set mrsAdvice = Nothing
    End If
    
End Sub

Private Sub optType_Click(Index As Integer)
    If Index = 1 Then
        lblStopNote.Caption = "路径表中用到的已停用项目"
        txtFind.ToolTipText = "查找下一个(F3)"
    Else
        lblStopNote.Caption = "路径表中在用的诊疗项目"
        txtFind.ToolTipText = "按编码、名称查找诊疗项目"
    End If
    If Me.Visible Then
        txtFind.Text = ""
        txtFind.SetFocus
        Call RefreshData
    End If
End Sub

Private Sub picAdvice_Resize()
    On Error Resume Next
    ucAdvice.Width = picAdvice.ScaleWidth - 90
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    cmdBatExe.Left = picBottom.ScaleWidth - (cmdBatExe.Width + cmdQuit.Width + 400)
    cmdQuit.Left = picBottom.ScaleWidth - (cmdQuit.Width + 300)
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    Line1(3).X1 = 0
    Line1(3).X2 = picInfo.ScaleWidth
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    fraFind.Move 120, 60, 4455, 900
    lblStopNote.Move 120, picTop.Height - 45
    With rptList
        .Left = 120
        .Top = picTop.Height + lblNote(0).Height - 45
        .Width = picLeft.ScaleWidth - .Left * 2
        .Height = picLeft.ScaleHeight - .Top + 15
    End With
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
   
    '路径清单
    lblNote(0).Move 0, 0, 1000, 255
    With rptPath
        .Left = 0
        .Top = lblNote(0).Top + lblNote(0).Height
        .Width = picMain.Width - 180
        .Height = picMain.Height - .Top - 100
    End With
End Sub

Private Sub InitRPTListColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    With rptList
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(COL_分类, "分类", 60, True)
    
        '分类，名称，编码，商品名，产地，药品剂型
        Set objCol = .Columns.Add(COL_编码, "编码", 80, True): objCol.Visible = True
        Set objCol = .Columns.Add(COL_名称, "名称", 200, False): objCol.Visible = True
        
        Set objCol = .Columns.Add(COL_商品名, "商品名", 100, True)
        Set objCol = .Columns.Add(COL_产地, "产地", 75, True)
        Set objCol = .Columns.Add(COL_药品剂型, "药品剂型", 75, True)
        '隐藏列
        Set objCol = .Columns.Add(COL_诊疗项目ID, "诊疗项目ID", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_收费细目ID, "收费细目ID", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_诊疗类别, "诊疗类别", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_操作类型, "操作类型", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_执行分类, "执行分类", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_计算方式, "计算方式", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_标本部位, "标本部位", 0, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_检查方法, "检查方法", 0, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_内容ID, "内容ID", 0, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_相关ID, "相关ID", 0, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_简码, "简码", 0, True): objCol.Visible = False
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = COL_分类
        Next
        
        rptList.Populate
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的诊疗项目..."
        End With
        
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False 'True 时单击列时,自动将列名添加到分组中

        
        .GroupsOrder.Add .Columns(COL_分类)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(COL_编码)
        .SortOrder(0).SortAscending = True
    End With
End Sub


Private Sub InitRPTPathColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    With rptPath
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(Path_ID, "ID", 0, True)
        objCol.Visible = False
        Set objCol = .Columns.Add(Path_选择, "选择", 60, True)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.Editable = True
        objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(Path_分类, "分类", 100, True)
        objCol.Visible = True
        Set objCol = .Columns.Add(Path_编码, "编码", 100, True)
        objCol.Visible = True
        Set objCol = .Columns.Add(Path_名称, "名称", 200, True)
        objCol.Visible = True
        Set objCol = .Columns.Add(Path_版本, "版本", 45, True)
        objCol.Visible = True
        Set objCol = .Columns.Add(Path_说明, "说明", 200, True)
        objCol.Visible = True
        
        For Each objCol In .Columns
            If objCol.Index <> Path_选择 Then
                objCol.Editable = False
            End If
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的临床路径..."
        End With
        
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
         .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(Path_分类)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
     
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(Path_编码)
        .SortOrder(0).SortAscending = True
    End With
End Sub

Private Sub LoadStopedItem()
'---------------------------------------
'功能:加载未停用路径表中用到的已停用项目或药品库存为零的项目
'参数:
'说明:
'---------------------------------------
    Dim strSql As String, str类别 As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strIDs As String
    Dim objRecord As ReportRecord
    Dim lng组ID As Long
    Dim str组Id As String
    Dim lngBegin As Long
    Dim strtxt As String
    
    Dim i As Long
    Dim j As Long
    If optType(1).Value Then
        strSql = "Select a.诊疗项目id, a.收费细目id, f.类别, Nvl(g.编码, f.编码) As 编码," & vbNewLine & _
                "       Nvl(g.名称, f.名称) || Decode(Nvl(g.规格, '0'), '0', '', ' ' || g.规格) As 名称, f.操作类型, f.执行分类, f.计算方式, f.计算单位, g.产地," & vbNewLine & _
                "       k.名称 As 商品名, h.药品剂型, -null As 标本部位, -null As 检查方法, -null As ID, -null As 相关id" & vbNewLine & _
                "From 诊疗项目目录 F, 收费项目目录 G, 药品特性 H, 收费项目别名 K," & vbNewLine & _
                "     (" & vbNewLine & _
                "     Select Distinct a.诊疗项目id, a.收费细目id" & vbNewLine & _
                "       From 路径医嘱内容 A, 诊疗项目目录 B, 收费项目目录 C, 临床路径医嘱 E, 临床路径项目 F, 临床路径版本 D" & vbNewLine & _
                "       Where a.诊疗项目id = b.Id And a.收费细目id = c.Id(+) And a.Id = e.医嘱内容id And e.路径项目id = f.Id And f.路径id = d.路径id And" & vbNewLine & _
                "             f.版本号 = d.版本号 And d.停用时间 Is Null And" & vbNewLine & _
                "             ((Nvl(b.撤档时间, Sysdate) <> To_Date('3000-01-01', 'yyyy-mm-dd') And b.撤档时间 Is Not Null) Or" & vbNewLine & _
                "             (Nvl(c.撤档时间, Sysdate) <> To_Date('3000-01-01', 'yyyy-mm-dd') And c.撤档时间 Is Not Null) Or Exists" & vbNewLine & _
                "              (Select 1" & vbNewLine & _
                "               From 药品库存" & vbNewLine & _
                "               Where 药品id = a.收费细目id And (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And 性质 = 1" & vbNewLine & _
                "               Group By 药品id" & vbNewLine & _
                "               Having Nvl(Sum(可用数量), 0) <= 0)) And b.类别 In ('5', '6', '7')" & vbNewLine & _
                "       ) A" & vbNewLine & _
                "Where a.诊疗项目id = f.Id And a.诊疗项目id = h.药名id(+) And a.收费细目id = g.Id(+) And a.收费细目id = k.收费细目id(+) And k.性质(+) = 3 And" & vbNewLine & _
                "      k.码类(+) = 1"
        strSql = strSql & " Union All "
        strSql = strSql & "Select a.诊疗项目id, -null As 收费细目id, a.类别, a.编码, a.名称, a.操作类型, a.执行分类, a.计算方式, '' As 计算单位, '' As 产地, '' As 商品名, '' As 药品剂型," & vbNewLine & _
                "       a.标本部位, a.检查方法, a.Id, a.相关id" & vbNewLine & _
                "From (Select h.路径项目id, h.诊疗项目id, h.编码, h.类别, b.名称, h.操作类型, h.执行分类, h.计算方式, a.Id, a.相关id, a.序号, a.标本部位, a.检查方法" & vbNewLine & _
                "       From (Select Distinct Nvl(a.相关id, a.Id) As 组id, b.路径项目id, a.诊疗项目id, g.类别, g.编码, g.名称 As 名称, g.操作类型, g.执行分类, g.计算方式" & vbNewLine & _
                "              From 路径医嘱内容 A, 临床路径医嘱 B, 临床路径项目 C, 临床路径版本 D, 诊疗项目目录 G" & vbNewLine & _
                "              Where a.Id = b.医嘱内容id And b.路径项目id = c.Id And c.路径id = d.路径id And c.版本号 = d.版本号 And d.停用人 Is Null And" & vbNewLine & _
                "                    a.诊疗项目id = g.Id And Nvl(g.撤档时间, Sysdate) <> To_Date('3000-01-01', 'yyyy-mm-dd') And g.撤档时间 Is Not Null And" & vbNewLine & _
                "                    g.类别 In ('D', 'C')) H, 路径医嘱内容 A, 诊疗项目目录 B" & vbNewLine & _
                "       Where (h.组id = a.Id Or h.组id = a.相关id) And a.诊疗项目id = b.Id" & vbNewLine & _
                "       Order By h.路径项目id, a.序号) A"
        strSql = strSql & " Union All "
        strSql = strSql & "Select Distinct a.诊疗项目id, -null As 收费细目id, a.类别, a.编码, a.名称, a.操作类型, a.执行分类,a.计算方式, '' As 计算单位, '' As 产地, '' As 商品名, '' As 药品剂型," & vbNewLine & _
                "                a.标本部位, a.检查方法, a.Id, a.相关id  " & vbNewLine & _
                "From (Select a.诊疗项目id, g.类别, g.编码, g.名称 As 名称, g.操作类型, g.执行分类,g.计算方式, a.标本部位, a.检查方法, -null As ID, -null As 相关id" & vbNewLine & _
                "       From 路径医嘱内容 A, 临床路径医嘱 B, 临床路径项目 C, 临床路径版本 D,诊疗项目目录 G" & vbNewLine & _
                "       Where a.Id = b.医嘱内容id And b.路径项目id = c.Id And c.路径id = d.路径id And c.版本号 = d.版本号 And d.停用人 Is Null And a.诊疗项目id = g.Id And" & vbNewLine & _
                "             Nvl(g.撤档时间, Sysdate) <> To_Date('3000-01-01', 'yyyy-mm-dd') And g.撤档时间 Is Not Null And" & vbNewLine & _
                "             g.类别 Not In ('D', 'C', '5', '6', '7')) A"
    Else
        strtxt = Trim(txtFind.Text)
        If strtxt = "" Then
            rptList.Records.DeleteAll
            rptList.Populate
            Exit Sub
        End If
        If zlCommFun.IsCharChinese(strtxt) Then
            strtxt = " And g.名称 like [1]"
        Else
            strtxt = " And g.编码 like [1]"
        End If
        strSql = "Select a.诊疗项目id, a.收费细目id, f.类别, f.编码, Nvl(g.名称, f.名称) || Decode(Nvl(g.规格, '0'), '0', '', ' ' || g.规格) As 名称, f.操作类型," & vbNewLine & _
            "       f.执行分类, f.计算方式, f.计算单位, g.产地, k.名称 As 商品名, h.药品剂型, -null As 标本部位, -null As 检查方法, -null As ID, -null As 相关id" & vbNewLine & _
            "From 诊疗项目目录 F, 收费项目目录 G, 药品特性 H, 收费项目别名 K," & vbNewLine & _
            "     (Select Distinct a.诊疗项目id, a.收费细目id" & vbNewLine & _
            "       From 路径医嘱内容 A, 临床路径医嘱 E, 临床路径项目 F, 临床路径版本 D, 诊疗项目目录 G" & vbNewLine & _
            "       Where a.诊疗项目id = G.Id And a.Id = e.医嘱内容id And e.路径项目id = f.Id And f.路径id = d.路径id And f.版本号 = d.版本号 And" & vbNewLine & _
            "             d.停用时间 Is Null And g.类别 In ('5', '6', '7')  " & strtxt & ") A" & vbNewLine & _
            "Where a.诊疗项目id = f.Id And a.诊疗项目id = h.药名id(+) And a.收费细目id = g.Id(+) And a.收费细目id = k.收费细目id(+) And k.性质(+) = 3 And" & vbNewLine & _
            "      k.码类(+) = 1"
        strSql = strSql & " Union All "
        strSql = strSql & "Select a.诊疗项目id, -null As 收费细目id, a.类别, a.编码, a.名称, a.操作类型, a.执行分类, a.计算方式, '' As 计算单位, '' As 产地, '' As 商品名, '' As 药品剂型," & vbNewLine & _
            "       a.标本部位, a.检查方法, a.Id, a.相关id" & vbNewLine & _
            "From (Select h.路径项目id, h.诊疗项目id, h.编码, h.类别, b.名称, h.操作类型, h.执行分类, h.计算方式, a.Id, a.相关id, a.序号, a.标本部位, a.检查方法" & vbNewLine & _
            "       From (Select Distinct Nvl(a.相关id, a.Id) As 组id, b.路径项目id, a.诊疗项目id, g.类别, g.编码, g.名称 As 名称, g.操作类型, g.执行分类, g.计算方式" & vbNewLine & _
            "              From 路径医嘱内容 A, 临床路径医嘱 B, 临床路径项目 C, 临床路径版本 D, 诊疗项目目录 G" & vbNewLine & _
            "              Where a.Id = b.医嘱内容id And b.路径项目id = c.Id And c.路径id = d.路径id And c.版本号 = d.版本号 And d.停用人 Is Null And" & vbNewLine & _
            "                    a.诊疗项目id = g.Id And g.类别 In ('D', 'C') " & strtxt & ") H, 路径医嘱内容 A, 诊疗项目目录 B" & vbNewLine & _
            "       Where (h.组id = a.Id Or h.组id = a.相关id) And a.诊疗项目id = b.Id" & vbNewLine & _
            "       Order By h.路径项目id, a.序号) A"
            strSql = strSql & " Union All "
        strSql = strSql & "Select Distinct a.诊疗项目id, -null As 收费细目id, a.类别, a.编码, a.名称, a.操作类型, a.执行分类, a.计算方式, '' As 计算单位, '' As 产地, '' As 商品名," & vbNewLine & _
                "                '' As 药品剂型, a.标本部位, a.检查方法, a.Id, a.相关id" & vbNewLine & _
                "From (Select a.诊疗项目id, g.类别, g.编码, g.名称 As 名称, g.操作类型, g.执行分类, g.计算方式, a.标本部位, a.检查方法, -null As ID, -null As 相关id" & vbNewLine & _
                "       From 路径医嘱内容 A, 临床路径医嘱 B, 临床路径项目 C, 临床路径版本 D, 诊疗项目目录 G" & vbNewLine & _
                "       Where a.Id = b.医嘱内容id And b.路径项目id = c.Id And c.路径id = d.路径id And c.版本号 = d.版本号 And d.停用人 Is Null And" & vbNewLine & _
                "             a.诊疗项目id = g.Id And g.类别 Not In ('D', 'C', '5', '6', '7') " & strtxt & ") A"
        
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(Trim(txtFind.Text)) & "%")
    txtFind.Text = ""
    With rptList
        For i = 1 To rsTmp.RecordCount
            If InStr(",5,6,7,", rsTmp!类别 & "") > 1 Then
                str类别 = "01-药品"
            ElseIf rsTmp!类别 & "" = "D" Then
                str类别 = "02-检查"
            ElseIf rsTmp!类别 & "" = "C" Then
                str类别 = "03-检验"
            Else
                str类别 = "04-其他"
            End If
            Set objRecord = .Records.Add
            objRecord.AddItem str类别
            objRecord.AddItem rsTmp!编码 & ""
            objRecord.AddItem rsTmp!名称 & ""
            objRecord.AddItem rsTmp!商品名 & ""
            objRecord.AddItem rsTmp!产地 & ""
            objRecord.AddItem rsTmp!药品剂型 & ""
            objRecord.AddItem rsTmp!诊疗项目ID & ""
            objRecord.AddItem rsTmp!收费细目ID & ""
            objRecord.AddItem rsTmp!类别 & ""
            objRecord.AddItem rsTmp!操作类型 & ""
            objRecord.AddItem rsTmp!执行分类 & ""
            objRecord.AddItem rsTmp!计算方式 & ""
            objRecord.AddItem rsTmp!标本部位 & ""
            objRecord.AddItem rsTmp!检查方法 & ""
            objRecord.AddItem rsTmp!ID & ""
            objRecord.AddItem rsTmp!相关id & ""
            objRecord.AddItem zlCommFun.SpellCode(NVL(rsTmp!名称) & "※0")
            rsTmp.MoveNext
        Next

        '隐藏行
        For i = 0 To .Records.count - 1
            str组Id = IIf(.Records.Record(i).Item(COL_相关ID).Value <> "", .Records.Record(i).Item(COL_相关ID).Value, .Records.Record(i).Item(COL_内容ID).Value)
            If .Records.Record(i).Item(COL_诊疗类别).Value = "D" Then
                If .Records.Record(i).Item(COL_内容ID).Value <> str组Id Then    '组ID
                    .Records.Record(i).Visible = False
                Else
                    .Records.Record(i).Visible = True
                    .Records.Record(i).Item(COL_名称).Value = AdviceMakeText(i, "D", strTmp)
                    .Records.Record(i).Tag = strTmp
                    strIDs = strIDs & "," & .Records.Record(i).Item(COL_诊疗项目ID).Value & "_" & i
                End If
            ElseIf .Records.Record(i).Item(COL_诊疗类别).Value = "C" Then
                If .Records.Record(i).Item(COL_内容ID).Value <> str组Id Then    '组ID
                    If lng组ID <> Val(str组Id) Then
                        lng组ID = str组Id
                        lngBegin = i    '记录检验行的首行
                    End If
                    .Records.Record(i).Visible = False
                Else
                    .Records.Record(i).Visible = True
                    .Records.Record(i).Item(COL_名称).Value = AdviceMakeText(i, "C", strTmp, lngBegin)
                    .Records.Record(i).Tag = strTmp
                    strIDs = strIDs & "," & .Records.Record(i).Item(COL_诊疗项目ID).Value & "_" & i
                End If
            End If
        Next
        strIDs = strIDs & ","
        
        For i = 0 To .Records.count - 1
            If .Records.Record(i).Visible And InStr(",C,D,", "," & .Records.Record(i).Item(COL_诊疗类别).Value & ",") > 0 Then
                strTmp = .Records.Record(i).Item(COL_诊疗项目ID).Value
                
                If InStr(Mid(strIDs, 1, InStr(strIDs, "," & strTmp & "_" & i)), "," & strTmp & "_") > 0 Then  '找到前面诊疗项目ID与当前诊疗项目相同的
                    For j = i - 1 To 0 Step -1
                        If .Records.Record(j).Visible And .Records.Record(i).Item(COL_诊疗类别).Value = .Records.Record(j).Item(COL_诊疗类别).Value Then
                            If CompareStr(.Records.Record(i).Tag, .Records.Record(j).Tag) Then  '整组内容相同
                                .Records.Record(j).Item(COL_内容ID).Value = .Records.Record(j).Item(COL_内容ID).Value & "," & .Records.Record(i).Item(COL_内容ID).Value
                                .Records.Record(i).Visible = False
                            End If
                        End If
                    Next
                End If
            End If
        Next
        .Populate
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub picSplit_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'功能:拖动停用项目列表
   If Button = 1 Then
        On Error Resume Next
        If picSplit.Left + X < Me.ScaleWidth / 10 Or picSplit.Left + X > Me.ScaleWidth / 10 * 9 Then Exit Sub
        picSplit.Left = picSplit.Left + X
        picLeft.Width = picLeft.Width + X
        picTop.Left = picTop.Left + X: picTop.Width = picTop.Width - X
        picMain.Left = picMain.Left + X: picMain.Width = picMain.Width - X
        picAdvice.Left = picAdvice.Left + X: picAdvice.Width = picAdvice.Width - X
        fraSplit.Left = fraSplit.Left + X: fraSplit.Width = fraSplit.Width - X
    End If
End Sub

Private Sub SetFilterInfo()
'功能:设置过滤信息
'
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim lng收费细目ID As Long
    Dim lng诊疗项目ID As Long
    Dim str类别 As String
    Dim blnTooLong As Boolean
    Dim str内容ID As String
    Dim strTmp As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    '清空
    Call ClearParaValue
    
    lng收费细目ID = Val(rptList.SelectedRows(0).Record(COL_收费细目ID).Value)
    lng诊疗项目ID = Val(rptList.SelectedRows(0).Record(COL_诊疗项目ID).Value)
    str类别 = rptList.SelectedRows(0).Record(COL_诊疗类别).Value
    str内容ID = "," & rptList.SelectedRows(0).Record(COL_内容ID).Value & ","
    '收费细目ID为空才取诊疗项目ID
    
    strSql = "Select d.单次用量, d.总给予量, d.执行频次, f.名称 As 用法,f.计算方式, f.Id as 诊疗项目ID  " & _
            " From 临床路径版本 A, 临床路径项目 B, 临床路径医嘱 C, 路径医嘱内容 D, 路径医嘱内容 E, 诊疗项目目录 F " & _
            " Where a.路径ID = b.路径id And a.版本号 = b.版本号 and A.停用人 is null And b.Id = c.路径项目id And c.医嘱内容id = d.Id And " & _
            IIf("K" = rptList.SelectedRows(0).Record(COL_诊疗类别).Value, "d.id = e.相关Id(+)", "d.相关id = e.Id(+)") & " And e.诊疗项目id = f.Id(+) "
    
    If str类别 = "C" Or str类别 = "D" Then
        If Len(str内容ID) > 4000 Then
            blnTooLong = True
        End If
        strSql = strSql & " And Instr([2],','||NVl(d.相关Id,d.ID)||',' ) >0 "
    End If
    
    If lng收费细目ID = 0 Then
        strSql = strSql & " And d.诊疗项目id = [1] And Rownum < 2 "
    Else
        strSql = strSql & " And d.收费细目ID = [1] And Rownum < 2 "
    End If
    
    If blnTooLong Then
        j = 1
        Do While j < Len(str内容ID)
            strTmp = Mid(str内容ID, j, 4000)
            i = InStrRev(strTmp, ",")
            strTmp = Mid(strTmp, 1, i)
            j = j + i - 1
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng收费细目ID = 0, lng诊疗项目ID, lng收费细目ID), strTmp)
            If rsTmp.RecordCount > 0 Then
                Exit Do
            End If
        Loop
        If rsTmp.RecordCount = 0 Then Exit Sub
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng收费细目ID = 0, lng诊疗项目ID, lng收费细目ID), str内容ID)
        If rsTmp.RecordCount = 0 Then Exit Sub
    End If
    
    With rptList.SelectedRows(0)
        If InStr(",5,6,C,", "," & .Record(COL_诊疗类别).Value & ",") > 0 Then
            txt单量.Text = FormatEx(NVL(rsTmp!单次用量), 4)
            txt总量.Text = FormatEx(NVL(rsTmp!总给予量), 4)
            txt用法.Text = rsTmp!用法 & "": txt用法.Tag = rsTmp!诊疗项目ID & ""
            txt频率.Text = rsTmp!执行频次 & ""
        ElseIf .Record(COL_诊疗类别).Value = "D" Then
            txt总量.Text = rsTmp!总给予量 & ""
            txt频率.Text = rsTmp!执行频次 & ""
        ElseIf InStr(",1,2,", "," & .Record(COL_计算方式).Value & ",") > 0 Then '1-计量，2-计时
            txt单量.Text = FormatEx(NVL(rsTmp!单次用量), 4)
            txt总量.Text = FormatEx(NVL(rsTmp!总给予量), 4)
            txt频率.Text = rsTmp!执行频次 & ""
        End If
    End With
    If lng收费细目ID = 0 Then
        strSql = "Select a.类别 As 类别id, a.计算方式, a.执行频率, a.计算单位, NULL as 住院单位 From 诊疗项目目录 A Where a.Id = [1]"
    Else
        strSql = "Select a.类别 As 类别id, a.计算方式, a.执行频率, a.计算单位, b.住院单位" & vbNewLine & _
                "From 诊疗项目目录 A, 药品规格 B" & vbNewLine & _
                "Where a.Id = [1] And b.药品id = [2]"

    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng诊疗项目ID, lng收费细目ID)
    If rsTmp.RecordCount = 0 Then Exit Sub
    '单量单位
    If txt总量.Text = "" Then '长嘱
        If InStr(",5,6,", rsTmp!类别ID) > 0 Or InStr(",1,2,", NVL(rsTmp!计算方式, 0)) > 0 Then
            lbl单量单位.Caption = NVL(rsTmp!计算单位)   '药品为剂量单位
        End If
    Else
        If InStr(",5,6,", rsTmp!类别ID) > 0 Or (NVL(rsTmp!执行频率, 0) = 0 And InStr(",1,2,", NVL(rsTmp!计算方式, 0)) > 0) Then
            lbl单量单位.Caption = NVL(rsTmp!计算单位)   '药品为剂量单位
        End If
    End If

    '总量单位
    If txt总量.Text <> "" Then '临嘱
        If InStr(",5,6,", rsTmp!类别ID) > 0 Then
            '中、西成药临嘱的总量单位就是住院单位
            lbl总量单位.Caption = rsTmp!住院单位 & ""
        ElseIf rsTmp!类别ID = "4" Then
            lbl总量单位.Caption = rsTmp!住院单位 & ""  '散装单位
        Else
            '其它临嘱要输入总量
            '如果为一次性或计次临嘱缺省总量为1
            If NVL(rsTmp!执行频率, 0) = 1 Or NVL(rsTmp!计算方式, 0) = 3 Then
               lbl总量单位.Caption = 1
            End If
            lbl总量单位.Caption = NVL(rsTmp!计算单位)
        End If
    End If
        
    mblnChange = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub rptList_SelectionChanged()
    Dim objItem As ReportRecordItem
    Dim strTmp As String
    
    cmdFind.Enabled = False
    If rptList.SelectedRows.count = 0 Then Exit Sub  '非正常情况
    If rptList.SelectedRows(0).GroupRow Then
        SetEditable -1, -1, -1, -1, -1, -1, -1
        Call ClearParaValue
        Call ClearPath
        Exit Sub
    End If
    With rptList.SelectedRows(0)
        '
        strTmp = "用法" '缺省设置为用法
        If InStr(",5,6,7,", "," & .Record(COL_诊疗类别).Value & ",") > 0 Then
            SetEditable 1, 1, 1, 1, 1, 1, 1
        ElseIf .Record(COL_诊疗类别).Value = "D" Then
            SetEditable -1, -1, -1, 1, -1, -1, 1
        ElseIf .Record(COL_诊疗类别).Value = "C" Then
            strTmp = "采集方式"
            SetEditable -1, -1, 1, 1, -1, 1, 1
        ElseIf InStr(",1,2,", "," & .Record(COL_计算方式).Value & ",") > 0 Then   '1-计量，2-计时
            SetEditable 1, 1, -1, 1, 1, -1, 1
        Else
            SetEditable -1, -1, -1, -1, -1, -1, -1
        End If
        lbl用法.Caption = strTmp
        If .Record(COL_诊疗类别).Value = "C" Then
            chkPra(chk_用法).Caption = "只替换采集相同的"
        Else
            chkPra(chk_用法).Caption = "只替换用法相同的"
        End If
        
        If mblnChange Or mstrSelItem <> .Record(COL_诊疗项目ID).Value & "_" & .Record(COL_收费细目ID).Value & "_" & .Record.Index Then
            '选择项目切换时,需要清空数据
            Call ClearPath
            mstrSelItem = .Record(COL_诊疗项目ID).Value & "_" & .Record(COL_收费细目ID).Value & "_" & .Record.Index
            Call SetFilterInfo
        End If
    End With
    cmdFind.Enabled = True
End Sub

Private Sub rptPath_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '如果点击表头的图片，就选中全部
    If Button = 1 Then
        If rptPath.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPath.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = Path_选择 Then
                    If rptPath.Columns(Path_选择).Icon = img16.ListImages("Check").Index - 1 Then
                        rptPath.Columns(Path_选择).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptPath.Records.count - 1
                            rptPath.Records(i)(Path_选择).Checked = False
                        Next
                    Else
                        rptPath.Columns(Path_选择).Icon = img16.ListImages("Check").Index - 1
                        For i = 0 To rptPath.Records.count - 1
                            rptPath.Records(i)(Path_选择).Checked = True
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPath_SortOrderChanged()
    Dim objCol As ReportColumn
    '排序时，强行先按分类排序
    '子项排序功能无效，它随主项一起排序
    If rptPath.SortOrder.count = 1 Then
        If rptPath.SortOrder(0).Index <> Path_分类 Then
            Set objCol = rptPath.SortOrder(0)
            rptPath.SortOrder.DeleteAll
            rptPath.SortOrder.Add rptPath.Columns(Path_分类)
            rptPath.SortOrder.Add objCol
        End If
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtFind.Text <> "" Then
        If optType(1).Value Then
            Call FindRPTList(True)
        Else
            Call RefreshData
        End If
        txtFind.SetFocus '定位到查找框
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And txtFind.Text <> "" Then
        Call FindRPTList(True)
        txtFind.SetFocus '定位到查找框
    End If
End Sub

Private Sub txt单量_Change()
    mblnChange = True
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub txt频率_Change()
    mblnChange = True
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub txt频率_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str范围 As String, int频率 As Integer, vRect As RECT
    Dim lng诊疗项目ID As Long
    
    If rptList.SelectedRows.count = 0 Then Exit Sub  '非正常情况
    With rptList.SelectedRows(0)
        If KeyAscii = 13 Then
            KeyAscii = 0
            If txt频率.Text = "" Then
                If cmd频率.Enabled And cmd频率.Visible Then cmd频率_Click
            Else
                int频率 = Get项目频率
                If txt总量.Text <> "" Then '临嘱
                    If .Record(COL_诊疗类别).Value <> "7" And int频率 = 0 Then
                        str范围 = "1,-1" '临嘱可以为一次性
                    Else
                        str范围 = Get频率范围(int频率)
                    End If
                Else
                    str范围 = Get频率范围(int频率)
                    int频率 = int频率 = Decode(str范围, "1", 0, "2", 0, "-1", 1, "-2", 2, "-3", 1, "-5", 1)
                End If
                
                '可选择频率的常用频率
                lng诊疗项目ID = Val(.Record(COL_诊疗项目ID).Value)
                strSql = ""
                If InStr("," & str范围 & ",", ",1,") > 0 Then
                    strSql = " And (Exists(Select 1 From 诊疗用法用量 Where 项目ID=[4] And 用法ID is NULL And 频次=A.编码 And A.适用范围=1)" & _
                        " Or (Select Count(*) From 诊疗用法用量 Where 项目ID=[4] And 用法ID is NULL And 频次 Is Not NULL)<=1)"
                End If
                strSql = _
                    " Select Rownum as ID,A.编码,A.名称,A.简码," & _
                    " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位,A.适用范围 as 范围ID" & _
                    " From 诊疗频率项目 A" & _
                    " Where (Instr([3],','||A.适用范围||',')>0   Or a.适用范围=[5])" & strSql & _
                    " And (A.编码 Like [1] Or Upper(A.名称) Like [2]" & _
                    " Or Upper(A.简码) Like [2] Or Upper(A.英文名称) Like [2])" & _
                    " Order by A.适用范围,A.编码"
                vRect = zlControl.GetControlRect(txt频率.Hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗频率", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt频率.Height, blnCancel, False, True, UCase(txt频率.Text) & "%", _
                    gstrLike & UCase(txt频率.Text) & "%", "," & str范围 & ",", lng诊疗项目ID, IIf(txt总量.Text <> "", -5, -3))
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到匹配的诊疗频率项目。", vbInformation, gstrSysName
                    End If
                    Call zlControl.TxtSelAll(txt频率)
                    txt频率.SetFocus: Exit Sub
                End If
                txt频率.Text = rsTmp!名称 & ""
                Call zlControl.TxtSelAll(txt频率)
                txt频率.SetFocus
            End If
        End If
    End With
End Sub

Private Sub txt用法_Change()
    mblnChange = True
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub txt用法_KeyPress(KeyAscii As Integer)
    Dim int类型 As Integer
    Dim strSql As String
    Dim strLike As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    
    If rptList.SelectedRows.count = 0 Then Exit Sub  '非正常情况
    If rptList.SelectedRows(0).GroupRow Then Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0
        With rptList.SelectedRows(0)
            If txt用法.Text = "" Then
                If cmd用法.Enabled And cmd用法.Visible Then cmd用法_Click
            Else
                If InStr(",5,6,", .Record(COL_诊疗类别).Value) > 0 Then
                    int类型 = 2 '给药途径
                ElseIf .Record(COL_诊疗类别).Value = "C" Then
                    int类型 = 6 '采集方法
                ElseIf .Record(COL_诊疗类别).Value = "K" Then
                    int类型 = 8 '输血途径
                Else
                    int类型 = 4 '中药用法
                End If
                If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
                    strSql = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[4] And 性质>0)" & _
                        " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                            " Where A.用法ID=B.ID And B.服务对象 IN([6],3) And A.项目ID=[4] And A.性质>0)<=1)"
                End If
                
                '优化
                strLike = gstrLike
                If Len(txt用法.Text) < 2 Then strLike = ""
                
                strSql = "Select Distinct A.ID,A.编码,A.名称" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B" & _
                    " Where A.ID=B.诊疗项目ID" & _
                    " And A.类别='E' And A.操作类型=[3] And A.服务对象 IN([6],3)" & strSql & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2])" & _
                    Decode(gint简码, 0, " And B.码类 IN([5],3)", 1, " And B.码类 IN([5],3)", "") & _
                    " Order by A.编码"
                vRect = zlControl.GetControlRect(txt用法.Hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, lbl用法.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, UCase(txt用法.Text) & "%", _
                    strLike & UCase(txt用法.Text) & "%", CStr(int类型), Val(.Record(COL_诊疗项目ID).Value), gint简码 + 1, 2)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到匹配的" & lbl用法.Caption & "。", vbInformation, gstrSysName
                    End If
                    Call zlControl.TxtSelAll(txt用法)
                    txt用法.SetFocus: Exit Sub
                End If
                txt用法.SetFocus
                txt用法.Text = rsTmp!名称 & ""
                txt用法.Tag = rsTmp!ID & ""
                Call zlControl.TxtSelAll(txt用法)
            End If
        End With
    End If
End Sub

Private Sub txt总量_Change()
    mblnChange = True
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub txt总量_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt单量_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function Get频率范围(ByVal lng频率性质 As Long) As Integer
    Dim lngFind As Long
    
    With rptList.SelectedRows(0)
        If .Record(COL_诊疗类别).Value = "7" Then
            Get频率范围 = 2 '中医
        Else
            If lng频率性质 = 0 Then
                Get频率范围 = 1 '可选频率的项目使用西医频率项目
            ElseIf lng频率性质 = 1 Then
                Get频率范围 = -1 '一次性
            ElseIf lng频率性质 = 2 Then
                Get频率范围 = -2 '持续性
            End If
        End If
    End With
End Function

Private Function Get项目频率() As Integer
'功能：获取指定项目的原始执行频率属性
'参数：lngRow=当前可见行
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 执行频率 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(rptList.SelectedRows(0).Record(COL_诊疗项目ID).Value))
    If Not rsTmp.EOF Then Get项目频率 = NVL(rsTmp!执行频率, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetEditable(Optional int单量 As Integer, Optional int总量 As Integer, _
    Optional int用法 As Integer, Optional int频率 As Integer, Optional intPra用量 As Integer, _
    Optional intPra用法 As Integer, Optional intPra频率 As Integer)
'功能：设置指定编辑项的可用状态
'参数：0-保持不变,-1-禁止,1-允许
    
    If int单量 = 1 Then
        txt单量.Enabled = True
        txt单量.BackColor = Color_Enabled
        lbl单量单位.Visible = True
    ElseIf int单量 = -1 Then
        txt单量.Enabled = False
        txt单量.BackColor = Color_UNEnabled
        lbl单量单位.Visible = False
    End If
    
    If int总量 = 1 Then
        txt总量.Enabled = True
        txt总量.BackColor = Color_Enabled
    ElseIf int总量 = -1 Then
        txt总量.Enabled = False
        txt总量.BackColor = Color_UNEnabled
    End If
    
    If int频率 = 1 Then
        txt频率.Enabled = True
        txt频率.BackColor = Color_Enabled
        cmd频率.Enabled = True
    ElseIf int频率 = -1 Then
        txt频率.Enabled = False
        cmd频率.Enabled = False
        txt频率.BackColor = Color_UNEnabled
    End If
    
    If int用法 = 1 Then
        cmd用法.Enabled = True
        txt用法.Enabled = True
        txt用法.BackColor = Color_Enabled
    ElseIf int用法 = -1 Then
        cmd用法.Enabled = False
        txt用法.Enabled = False
        txt用法.BackColor = Color_UNEnabled
    End If
    
    If intPra用量 = 1 Then
        chkPra(chk_用量).Enabled = True
        chkPra(chk_用量).Value = Checked
    ElseIf intPra用量 = -1 Then
        chkPra(chk_用量).Enabled = False
        chkPra(chk_用量).Value = Unchecked
    End If
    
    If intPra用法 = 1 Then
        chkPra(chk_用法).Enabled = True
        chkPra(chk_用法).Value = Checked
    ElseIf intPra用法 = -1 Then
        chkPra(chk_用法).Enabled = False
        chkPra(chk_用法).Value = Unchecked
    End If
    
    If intPra频率 = 1 Then
        chkPra(chk_频率).Enabled = True
        chkPra(chk_频率).Value = Checked
    ElseIf intPra频率 = -1 Then
        chkPra(chk_频率).Enabled = False
        chkPra(chk_频率).Value = Unchecked
    End If
    
End Sub

Private Function ShowAdvice() As Boolean
'功能：显示路径项目对应的医嘱内容(临床路径项目编辑)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Long, j As Long
    
    strSql = ""
    '生成动态SQL
    With mrsAdvice
        .Filter = ""
        Do While Not .EOF
            strSql = strSql & " Union ALL Select "
            For i = 0 To .Fields.count - 1
                If Not IsNull(.Fields(i).Value) Then
                    If Rec.IsType(.Fields(i).Type, adVarChar) Then
                        strSql = strSql & "'" & Replace(Replace(.Fields(i).Value, "[", "("), "]", ")") & "'"
                    Else
                        strSql = strSql & .Fields(i).Value '没有日期型
                    End If
                Else
                    If Rec.IsType(.Fields(i).Type, adBigInt) Or Rec.IsType(.Fields(i).Type, adSmallInt) Or Rec.IsType(.Fields(i).Type, adSingle) Then
                        strSql = strSql & "-Null"
                    Else
                        strSql = strSql & "Null"
                    End If
                End If
                strSql = strSql & " As " & .Fields(i).Name & ","
            Next
            strSql = Left(strSql, Len(strSql) - 1) & " From Dual"
            .MoveNext
        Loop
        .Filter = ""
        strSql = Mid(strSql, 12)
    End With
    
    If strSql = "" Then
        Call ucAdvice.ShowAdvice(4, "", 0, 0, 0)
    Else
        Call ucAdvice.ShowAdvice(4, strSql, 0, 0, 0)
    End If
    ShowAdvice = True
End Function

Private Sub InitAdviceRecordset()
    If Not mrsAdvice Is Nothing Then
        If mrsAdvice.State = 1 Then mrsAdvice.Close
    End If
    Set mrsAdvice = New ADODB.Recordset
    
    mrsAdvice.Fields.Append "ID", adBigInt
    mrsAdvice.Fields.Append "是否缺省", adSmallInt
    mrsAdvice.Fields.Append "是否备选", adSmallInt
    mrsAdvice.Fields.Append "相关ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "序号", adBigInt
    mrsAdvice.Fields.Append "期效", adSmallInt
    mrsAdvice.Fields.Append "诊疗项目ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "医嘱内容", adVarChar, 1000, adFldIsNullable
    mrsAdvice.Fields.Append "单次用量", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "总给予量", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "标本部位", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "检查方法", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "医生嘱托", adVarChar, 1000, adFldIsNullable
    mrsAdvice.Fields.Append "执行频次", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "频率次数", adSmallInt, , adFldIsNullable
    mrsAdvice.Fields.Append "频率间隔", adSmallInt, , adFldIsNullable
    mrsAdvice.Fields.Append "间隔单位", adVarChar, 10, adFldIsNullable
    mrsAdvice.Fields.Append "执行性质", adSmallInt
    mrsAdvice.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "时间方案", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "配方ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "组合项目ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "执行标记", adSingle, , adFldIsNullable
    
    mrsAdvice.CursorLocation = adUseClient
    mrsAdvice.LockType = adLockOptimistic
    mrsAdvice.CursorType = adOpenStatic
    mrsAdvice.Open
End Sub

Private Sub InitSchemeRecordset(rsScheme As ADODB.Recordset)
    Set rsScheme = New ADODB.Recordset
    rsScheme.Fields.Append "是否备选", adSmallInt
    rsScheme.Fields.Append "是否缺省", adSmallInt
    rsScheme.Fields.Append "序号", adBigInt
    rsScheme.Fields.Append "相关序号", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "期效", adSmallInt
    rsScheme.Fields.Append "诊疗项目ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "医嘱内容", adVarChar, 1000, adFldIsNullable
    rsScheme.Fields.Append "天数", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "单次用量", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "总给予量", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "医生嘱托", adVarChar, 1000, adFldIsNullable
    rsScheme.Fields.Append "执行频次", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "频率次数", adSmallInt, , adFldIsNullable
    rsScheme.Fields.Append "频率间隔", adSmallInt, , adFldIsNullable
    rsScheme.Fields.Append "间隔单位", adVarChar, 10, adFldIsNullable
    rsScheme.Fields.Append "时间方案", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "执行性质", adSmallInt
    rsScheme.Fields.Append "标本部位", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "检查方法", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "配方ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "组合项目ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "执行标记", adSingle, , adFldIsNullable
    
    rsScheme.CursorLocation = adUseClient
    rsScheme.LockType = adLockOptimistic
    rsScheme.CursorType = adOpenStatic
    rsScheme.Open
End Sub

Private Function SaveData(ByVal strPath As String, ByVal strPathTag As String) As Boolean
'功能:完成批量替换
'参数:
'      strPath    如:路径ID1,版本号1:路径ID2,版本号2:....
'      strPathTag 形如: 路径Id_最新版本1,路径Id_最新版本2:...
    Dim str组IDs As String, lng序号 As Long
    Dim lng诊疗项目ID As Long, lng收费细目ID As Long
    Dim str类别 As String
    Dim strSql As String
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim arrSQL As Variant
    Dim blnTran As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strItemIDs As String  '记录路径项目ID
    Dim strItemAdvices As String   '行如:路径项目ID1,组医嘱ID1:路径项目ID2,组医嘱ID2...
    Dim colItem As Collection
    Dim colAdvice As Collection
    
    On Error GoTo errH
    
    With rptList.SelectedRows(0)
        lng诊疗项目ID = Val(.Record(COL_诊疗项目ID).Value)
        lng收费细目ID = Val(.Record(COL_收费细目ID).Value)
        str类别 = .Record(COL_诊疗类别).Value
        '找到替换的项目
        If InStr(",5,6,", str类别) > 0 Then
            strSql = "Select *" & vbNewLine & _
            "From (Select Distinct /*+cardinality(A,10)*/ a.C1 as 路径ID,a.C2 as 版本号,c.路径项目ID,e.期效,e.Id As 内容id, E.相关id, E.诊疗项目id, E.收费细目id, E.序号 " & vbNewLine & _
            "From Table(f_Num2list2([1], ':', ',')) A, 临床路径项目 B, 临床路径医嘱 C, 路径医嘱内容 D, 路径医嘱内容 E" & vbNewLine & _
            "Where a.C1 = b.路径id And a.C2 = b.版本号 And b.Id = c.路径项目id And c.医嘱内容id = d.Id " & vbNewLine & _
            IIf(lng收费细目ID = 0, " And d.诊疗项目ID =[2]", " And d.收费细目ID =[2]") & _
            " And (d.相关id = e.相关id Or d.相关id = e.Id)"
                
            If chkPra(chk_用量).Value = vbChecked Then
                If txt单量.Text <> "" Then
                    strSql = strSql & " And d.单次用量 =[3] "
                End If
                If txt总量.Text <> "" Then
                    strSql = strSql & " And d.总给予量 = [4] "
                End If
            End If
            
            If chkPra(chk_用法).Value = vbChecked Then
                If txt用法.Text <> "" Then
                    strSql = strSql & " and exists (select 1 from 路径医嘱内容  H where H.id=d.相关id and H.诊疗项目id = [5]) "
                End If
            End If
            
            If chkPra(chk_频率).Value = vbChecked Then
                If txt频率.Text <> "" Then
                    strSql = strSql & " And d.执行频次 =[6] "
                End If
            End If
            strSql = strSql & ") 　order By 路径项目id, 序号"
        ElseIf InStr(",D,C,", str类别) > 0 Then
            '获取检查的医嘱ID
            str组IDs = IIf(.Record(COL_相关ID).Value = "", .Record(COL_内容ID).Value, .Record(COL_相关ID).Value)
            strSql = "Select /* +Rule */" & vbNewLine & _
                " a.C1 as  路径ID,a.C2 as 版本号,c.路径项目id, c.医嘱内容id as 内容ID" & vbNewLine & _
                "From Table(f_Num2list2([1], ':', ',')) A, 临床路径项目 B, 临床路径医嘱 C" & vbNewLine & _
                "Where a.C1 = b.路径id And a.C2 = b.版本号 And b.Id = c.路径项目id And Instr([7], ','||c.医嘱内容id||',') > 0"

        Else '其他类别 单条医嘱,手术医嘱（单条替换）
            strSql = "Select /*+ RULE*/" & vbNewLine & _
                    "a.C1 as  路径ID,a.C2 as 版本号,c.路径项目ID,d.Id As 内容id" & vbNewLine & _
                    "From Table(f_Num2list2([1], ':', ',')) A, 临床路径项目 B, 临床路径医嘱 C, 路径医嘱内容 D" & vbNewLine & _
                    "Where a.C1 = b.路径id And a.C2 = b.版本号 And b.Id = c.路径项目id And c.医嘱内容id = d.Id " & _
                    IIf(lng收费细目ID = 0, " And d.诊疗项目ID =[2]", " And d.收费细目ID =[2]")
            If chkPra(chk_用量).Value = vbChecked And InStr(",1,2,", "," & rptList.SelectedRows(0).Record(COL_计算方式).Value & ",") > 0 Then   '计量计时
               If txt单量.Text <> "" Then
                   strSql = strSql & " And d.单次用量 =[3] "
               End If
               If txt总量.Text <> "" Then
                   strSql = strSql & " And d.总给予量 = [4] "
               End If
            End If
            
            If chkPra(chk_频率).Value = vbChecked Then
                If txt频率.Text <> "" Then
                    strSql = strSql & " And d.执行频次 =[6] "
                End If
            End If
        End If
  
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPath, IIf(lng收费细目ID = 0, lng诊疗项目ID, lng收费细目ID), _
            Val(txt单量.Text), Val(txt总量.Text), Val(txt用法.Tag), txt频率.Text, "," & str组IDs & ",")
        If rsTmp.RecordCount = 0 Then
            MsgBox "没有找到满足条件的停用项目,替换失败!", vbInformation + vbOKOnly, "批量替换"
            Exit Function
        End If
        '只允许最新版本（最近一次审核版本）产生变动记录
        Set colItem = New Collection: Set colAdvice = New Collection
        strItemIDs = "": strItemAdvices = ""
        For i = 1 To rsTmp.RecordCount
            If Not InStr(strItemIDs & ",", "," & rsTmp!路径项目ID & ",") > 0 And InStr(strPathTag, rsTmp!路径ID & "_" & rsTmp!版本号) > 0 Then  '记下路径医嘱变动的项目ID
                If Len(strItemIDs & "," & rsTmp!路径项目ID) > 4000 Then
                    colItem.Add Mid(strItemIDs, 2)
                    strItemIDs = "," & rsTmp!路径项目ID
                Else
                    strItemIDs = strItemIDs & "," & rsTmp!路径项目ID
                End If
            End If
            If str类别 = "D" Or str类别 = "C" Then
                If Len(strItemAdvices & ":" & rsTmp!路径项目ID & "," & rsTmp!内容ID) > 4000 Then
                    colAdvice.Add Mid(strItemAdvices, 2)
                    strItemAdvices = ":" & rsTmp!路径项目ID & "," & rsTmp!内容ID
                Else
                    strItemAdvices = strItemAdvices & ":" & rsTmp!路径项目ID & "," & rsTmp!内容ID
                End If
            End If
            rsTmp.MoveNext
        Next
        If strItemIDs <> "" Then
            colItem.Add Mid(strItemIDs, 2)
        End If
        If strItemAdvices <> "" Then
            colAdvice.Add Mid(strItemAdvices, 2)
        End If
    End With
    arrSQL = Array()
    With mrsAdvice
        For i = 1 To colItem.count
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_路径医嘱变动_Insert('" & colItem(i) & "'," & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & UserInfo.姓名 & "')"
        Next
        rsTmp.MoveFirst
        If InStr(",5,6,", str类别) > 0 Then
            For i = 1 To rsTmp.RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If rsTmp!相关id & "" = "" Then
                    .Filter = "相关Id=NULL"
                Else
                    .Filter = "相关Id<>NULL"
                End If
                If (InStr(",5,6,", str类别) > 0 And lng诊疗项目ID = Val(rsTmp!诊疗项目ID & "")) Then
                    arrSQL(UBound(arrSQL)) = "zl_路径医嘱内容_Update(1," & rsTmp!内容ID & "," & ZVal(NVL(!诊疗项目ID)) & "," & ZVal(NVL(!收费细目ID)) & ",'" & !医嘱内容 & "'," & _
                    ZVal(NVL(!单次用量)) & "," & IIf(rsTmp!期效 = 1, ZVal(NVL(!总给予量)), "NULL") & ",'" & !标本部位 & "','" & !检查方法 & "','" & !医生嘱托 & "','" & !执行频次 & "'," & _
                    ZVal(NVL(!频率次数)) & "," & ZVal(NVL(!频率间隔)) & ",'" & !间隔单位 & "','" & !时间方案 & "')"
                Else
                    If rsTmp!相关id & "" = "" Then
                    '用法
                        arrSQL(UBound(arrSQL)) = "zl_路径医嘱内容_Update( 1," & rsTmp!内容ID & "," & ZVal(NVL(!诊疗项目ID)) & ", NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & !执行频次 & "'," & _
                                     ZVal(NVL(!频率次数)) & "," & ZVal(NVL(!频率间隔)) & ",'" & !间隔单位 & "','" & !时间方案 & "')"
                    Else
                    '一并给药,替换一并项目
                        arrSQL(UBound(arrSQL)) = "zl_路径医嘱内容_Update(1," & rsTmp!内容ID & ",Null, NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & !执行频次 & "'," & _
                                     ZVal(NVL(!频率次数)) & "," & ZVal(NVL(!频率间隔)) & ",'" & !间隔单位 & "','" & !时间方案 & "')"
                    End If
                End If
            
                rsTmp.MoveNext
            Next
        ElseIf InStr(",D,C,", str类别) > 0 Then
            '将检查检验医嘱先插入[路径医嘱内容]
            strTmp = "": .MoveFirst
            For i = 1 To .RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                lng序号 = lng序号 + 1
                arrSQL(UBound(arrSQL)) = "Zl_路径医嘱内容_Insert(" & _
                        !ID & "," & ZVal(NVL(!相关id, 0)) & "," & !序号 & "," & !期效 & "," & _
                        ZVal(NVL(!诊疗项目ID, 0)) & ",'" & NVL(!医嘱内容) & "'," & ZVal(NVL(!单次用量, 0)) & "," & _
                        ZVal(NVL(!总给予量, 0)) & "," & ZVal(NVL(!收费细目ID, 0)) & ",'" & NVL(!标本部位) & "'," & _
                        "'" & NVL(!检查方法) & "','" & NVL(!执行频次) & "'," & ZVal(NVL(!频率次数, 0)) & "," & _
                        ZVal(NVL(!频率间隔, 0)) & ",'" & NVL(!间隔单位) & "','" & NVL(!医生嘱托) & "'," & _
                        NVL(!执行性质, 0) & "," & ZVal(NVL(!执行科室ID, 0)) & ",'" & NVL(!时间方案) & "',Null,Null)"
                strTmp = strTmp & "," & !ID
                .MoveNext
            Next
            strTmp = Mid(strTmp, 2)
            '完成原医嘱的删除,新医嘱Id与路径项目ID的关联及该路径项目内所有医嘱序号重整
            For i = 1 To colAdvice.count
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_路径医嘱内容_Update(2,Null,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,Null,Null,Null,null,null,'" & strTmp & "','" & colAdvice(i) & "','" & str类别 & "')"
            Next
        Else
            '其他单条替换
            '输血K,输氧L,H护理,手术F,皮试等只替换诊疗项目Id;
            '计量计时
            '替换中药配方中某一药:只替换诊疗项目ID,收费细目ID,单次用量，但频率不改变
            For i = 1 To rsTmp.RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If str类别 = "7" Then
                    arrSQL(UBound(arrSQL)) = "zl_路径医嘱内容_Update(1," & rsTmp!内容ID & "," & ZVal(NVL(!诊疗项目ID)) & "," & ZVal(NVL(!收费细目ID)) & ",'" & !医嘱内容 & "'," & _
                                ZVal(NVL(!单次用量)) & ")"
                ElseIf InStr(",1,2,", "," & rptList.SelectedRows(0).Record(COL_计算方式).Value & ",") > 0 Then '计量计时
                    arrSQL(UBound(arrSQL)) = "zl_路径医嘱内容_Update(1," & rsTmp!内容ID & "," & ZVal(NVL(!诊疗项目ID)) & ",NULL,NULL," & ZVal(NVL(!单次用量)) & "," & ZVal(NVL(!总给予量)) & ")"
                Else
                    arrSQL(UBound(arrSQL)) = "zl_路径医嘱内容_Update(1," & rsTmp!内容ID & "," & ZVal(NVL(!诊疗项目ID)) & ")"
                End If
                rsTmp.MoveNext
            Next
        End If
    End With

    '提交数据
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        End If
    Next
    gcnOracle.CommitTrans: blnTran = False

    '替换成功，需要刷新界面数据
    SaveData = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdviceMakeText(ByVal lngRow As Long, ByVal str类别 As String, ByRef strTag As String, Optional ByVal lngBegin As Long) As String
'功能:构造检验，检查显示行内容
    Dim str部位 As String
    Dim str方法 As String
    Dim str部位Last As String
    Dim strReturn As String
    Dim str检验 As String, str标本  As String
    Dim i As Long
    
    str部位 = "": str方法 = "": strTag = ""
    With rptList.Records
        If str类别 = "D" Then
            For i = lngRow + 1 To .count - 1
                If Val(.Record(i).Item(COL_相关ID).Value) = Val(.Record(lngRow).Item(COL_内容ID).Value) Then
                    If .Record(i).Item(COL_标本部位).Value <> "" Then
                        If .Record(i).Item(COL_标本部位).Value <> str部位Last And str部位Last <> "" Then
                            str部位 = str部位 & "," & str部位Last & IIf(str方法 <> "", "(" & Mid(str方法, 2) & ")", "")
                            str方法 = ""
                        End If
                        
                        If .Record(i).Item(COL_检查方法).Value <> "" Then
                            str方法 = str方法 & "," & .Record(i).Item(COL_检查方法).Value
                        End If
                        
                        str部位Last = .Record(i).Item(COL_标本部位).Value
                        
                        '检查方法,标本部位
                        strTag = strTag & "," & .Record(i).Item(COL_标本部位).Value & "_" & .Record(i).Item(COL_检查方法).Value
                        
                    End If
                Else
                    Exit For
                End If
            Next
            If str部位Last <> "" Then
                str部位 = str部位 & "," & str部位Last & IIf(str方法 <> "", "(" & Mid(str方法, 2) & ")", "")
            End If
            str部位 = Mid(str部位, 2) '检查组合项目的部位
            strReturn = .Record(lngRow).Item(COL_名称).Value & ":" & str部位
        ElseIf str类别 = "C" Then
            str检验 = "": str标本 = ""
            
            For i = lngBegin To lngRow - 1    '检验数据未提采集方式
         
                str检验 = .Record(i).Item(COL_名称).Value & "," & str检验
                str标本 = .Record(i).Item(COL_标本部位).Value
                '记录检验标本部位
                strTag = strTag & "," & .Record(i).Item(COL_名称).Value & "_" & .Record(i).Item(COL_标本部位).Value
          
            Next
            str检验 = Left(str检验, Len(str检验) - 1)
            strReturn = str检验 & "(" & str标本 & ")"
        End If
    End With
    strTag = Mid(strTag, 2)
    If strTag = "" Then
        strTag = strReturn
    End If
    AdviceMakeText = strReturn
End Function

Private Sub ClearPath()
'功能:清除路径表数据
    rptPath.Records.DeleteAll
    rptPath.Populate
    Call InitAdviceTable
End Sub

Private Sub InitAdviceTable()
'功能:清除医嘱数据
    cmdEdit.Enabled = False
    cmdBatExe.Enabled = False
    Call InitAdviceRecordset
    Call ShowAdvice
End Sub

Private Sub RefreshData()
'功能:刷新数据
    rptList.Records.DeleteAll
    Call ClearParaValue
    Call ClearPath
    Call LoadStopedItem
    Call SetEditable(-1, -1, -1, -1, -1, -1, -1)
End Sub

Private Sub FindRPTList(Optional ByVal blnNext As Boolean)
'参数：blnNext=是否查找下一个
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    Call zlControl.TxtSelAll(txtFind)

    '开始查找行
    If rptList.SelectedRows.count > 0 Then blnHave = True
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0    'ReportControl的索引从是0开始
    Else
        i = rptList.SelectedRows(0).Index + 1
    End If

    
    For i = i To rptList.Rows.count - 1
        With rptList.Rows(i)
            If Not .GroupRow Then
                If IsNumeric(txtFind.Text) Then
                    '1X.输入全是数字时只匹配编码
                    If .Record(COL_编码).Value Like "*" & UCase(Trim(txtFind.Text)) & "*" Then
                        Exit For
                    End If
                ElseIf zlCommFun.IsCharAlpha(txtFind.Text) Then
                    'X1.输入全是字母时只匹配简码
                    If .Record(COL_简码).Value Like "*" & UCase(Trim(txtFind.Text)) & "*" Then
                        Exit For
                    End If
                ElseIf zlCommFun.IsCharChinese(txtFind.Text) Then
                    '包含汉字,则只匹配名称
                    If .Record(COL_名称).Value Like "*" & UCase(Trim(txtFind.Text)) & "*" Then
                        Exit For
                    End If
                End If
            End If
        End With
    Next
    
    
    If i <= rptList.Rows.count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set rptList.FocusedRow = rptList.Rows(i)
        If rptList.Visible Then rptList.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的诊疗项目。", vbInformation, gstrSysName
    End If
End Sub

Private Function CompareStr(ByVal str1 As String, ByVal str2 As String, Optional ByVal strDelimiter As String = ",") As Boolean
'功能:比较两个以逗号分隔的字符串是否完全相等,忽略字符串中的顺序
'参数:
'     str1-字符串1（分界符之间的字符不能出现重复的）
'     str2-字符串2
'     strDelimiter-分隔符
'返回值:True-相当,false-不想当
'说明: str1="1,2,3";str2="1,3,2" ,返回值 -true
    Dim arrOne As Variant
    Dim arrTwo As Variant
    Dim i As Long
    
    arrOne = Split(str1, strDelimiter)
    arrTwo = Split(str2, strDelimiter)
    
    str2 = strDelimiter & str2 & strDelimiter
    If UBound(arrOne) <> UBound(arrTwo) Then Exit Function
    
    For i = LBound(arrOne) To UBound(arrOne)
        If InStr(str2, strDelimiter & arrOne(i) & strDelimiter) = 0 Then
            Exit Function
        End If
    Next
    CompareStr = True
End Function

Private Sub ClearParaValue()
'功能:清空过滤参数值
    
    txt单量.Text = ""
    txt总量.Text = ""
    txt用法.Text = ""
    txt用法.Tag = ""
    txt频率.Text = ""
    lbl单量单位.Caption = ""
    lbl总量单位.Caption = ""
End Sub
