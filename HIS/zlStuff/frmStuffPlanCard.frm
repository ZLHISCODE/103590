VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmStuffPlanCard 
   Caption         =   "卫材采购计划"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmStuffPlanCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh生产商 
      Height          =   2325
      Left            =   240
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   4101
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf供应商选择 
      Height          =   2565
      Left            =   6360
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   8
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   4
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   5
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   9
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9930
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   1
         Top             =   950
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   3
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lbl编制方法 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编制方法:"
         Height          =   180
         Left            =   8070
         TabIndex        =   24
         Top             =   660
         Width           =   810
      End
      Begin VB.Label txt编制方法 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "临近期间平均参照法"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9000
         TabIndex        =   23
         Top             =   660
         Width           =   2355
      End
      Begin VB.Label txt计划类型 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1080
         TabIndex        =   22
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "金额合计："
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   3840
         Width           =   900
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   19
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   18
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   17
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   16
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   15
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "卫材采购计划单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label Lbl计划类型 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "计划类型:"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   660
         Width           =   810
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编制人"
         Height          =   180
         Left            =   300
         TabIndex        =   13
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编制日期"
         Height          =   180
         Left            =   2160
         TabIndex        =   12
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   7365
         TabIndex        =   11
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   9240
         TabIndex        =   10
         Top             =   4500
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1000
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffPlanCard.frx":22EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffPlanCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffPlanCard.frx":3080
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCode 
      Caption         =   "材料"
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmStuffPlanCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnFirst As Boolean                '第一次显示
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mintUnit As Integer            '0-散装单位,1-包装单位
Private mbln下限 As Boolean                 '仅提取低于储备下限的药品
Private mint上限 As Integer
Private mint下限 As Integer

Private mlng计划ID As Long
Private mlng库房id As Long
Private mint计划类型 As Integer
Private mint编制方法 As Integer
Private mstr供货商ID As String      '以id分隔
Private mbln中标单位 As Boolean '包含中标供货商,要与mstr供货单位一起启作用.
Private mstr期间  As String                  '月以六位表示,季以五位表示,年以四位表示
Dim mstrPrivs As String                     '权限
Private Const mlngModule = 1724
Private mint校验方式 As Integer     '0-不检查；1－提醒；2－禁止
Private mblnCheck As Boolean
Private mblnFirstCheck As Boolean
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看
Private mbln计划数量 As Boolean         'true-产生计划数量 false-不产生计划数量
Private mstrNow As String               '记录当前日期
Private Const mstrCaption As String = "卫材采购计划"
Private mstr重复卫材 As String '记录重复的卫材
Private mblnStart As Boolean

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

'=========================================================================================
Private Enum mHeadCol
    序号 = 1
    校验 = 2
    材料 = 3
    规格 = 4
    产地 = 5
    单位 = 6
    比例系数 = 7
    中标材料 = 8
    存储下限 = 9
    存储上限 = 10
    前期数量 = 11
    上期数量 = 12
    库存数量 = 13
    上期销量 = 14
    本期销量 = 15
    计划数量 = 16
    单价 = 17
    金额 = 18
    上次供应商 = 19
End Enum

Private Const mconIntColS  As Integer = 20     '总列数

Private Function CheckQualifications() As Boolean
    '根据参数校验卫材，生产商，供应商信息和资质效期
    Dim dateCurrent As Date
    Dim strCheck As String
    Dim strCheck_卫材 As String
    Dim strCheck_生产商 As String
    Dim strCheck_供应商 As String
    Dim intCheckType As Integer
    Dim arrColumn
    Dim rsTmp As ADODB.Recordset
    Dim intRow As Integer
    Dim strTmp_卫材 As String
    Dim strTmp_生产商 As String
    Dim strTmp_供应商 As String
    Dim strMsg_卫材 As String
    Dim strMsg_生产商 As String
    Dim strMsg_供应商 As String
    Dim intCount As Integer
    Dim blnFlag As Boolean
    Dim n As Integer
    Dim strMsgInfo As String
    Dim str生产商列表 As String
    Dim str供应商列表 As String
    Dim intCount_卫材 As Integer
    Dim intCount_生产商 As Integer
    Dim intCount_供应商 As Integer

'    On Error Resume Next
    On Error GoTo ErrHandle
    '资质校验项目和方式的保存格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
    strCheck = zlDatabase.GetPara("资质校验", glngSys, mlngModule, "")

    If InStr(1, strCheck, "|") = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    '取校验方式：0-不检查；1－提醒；2－禁止
    intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))

    If intCheckType = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    '取校验内容：
    strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)

    If strCheck = "" Then
        CheckQualifications = True
        Exit Function
    End If

    '分别取卫材，生产商，供应商需要校验的内容
    strCheck = strCheck & ";"
    arrColumn = Split(strCheck, ";")
    For n = 0 To UBound(arrColumn)
        If arrColumn(n) <> "" Then
            If Split(arrColumn(n), ",")(0) = "卫材" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_卫材 = IIf(strCheck_卫材 = "", "", strCheck_卫材 & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "卫材生产商" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_生产商 = IIf(strCheck_生产商 = "", "", strCheck_生产商 & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "卫材供应商" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_供应商 = IIf(strCheck_供应商 = "", "", strCheck_供应商 & ";") & Split(arrColumn(n), ",")(1)
            End If
        End If
    Next

    dateCurrent = CDate(Format(sys.Currentdate, "yyyy-mm-dd"))

    '分别校验卫材，生产商，供应商
    With mshBill
        .Redraw = False
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                If strCheck_卫材 <> "" Then
                    gstrSQL = "Select A.许可证号, A.许可证有效期 " & _
                              "From 材料特性 A " & _
                              "Where A.材料ID = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "校验卫材资质", Val(.TextMatrix(intRow, 0)))
                    
                    strTmp_卫材 = ""
'                    strMsg_卫材 = ""
                    blnFlag = False
                    
                    If Not rsTmp.EOF Then
                        If zlStr.Nvl(rsTmp!许可证号) = "" And InStr(strCheck_卫材, "许可证号") > 0 Then
                            strTmp_卫材 = .TextMatrix(intRow, mHeadCol.材料) & "：" & "无许可证号"
                            blnFlag = True
                        End If
                        
                        If zlStr.Nvl(rsTmp!许可证有效期) <> "" Then
                            If DateDiff("d", rsTmp!许可证有效期, dateCurrent) > 0 And InStr(strCheck_卫材, "许可证号") > 0 Then
                                strTmp_卫材 = IIf(strTmp_卫材 = "", .TextMatrix(intRow, mHeadCol.材料) & "：", strTmp_卫材 & ",") & "许可证已过期"
                            blnFlag = True
                            End If
                        End If
                    End If
    
                    If strTmp_卫材 <> "" Then
                        If intCount_卫材 <= 5 Then
                            strMsg_卫材 = strMsg_卫材 & strTmp_卫材 & vbCrLf
                        End If
                        intCount_卫材 = intCount_卫材 + 1
                    End If
                    If blnFlag = True Then SetBilCheckFlag intRow, mHeadCol.材料, False
                End If
                
                If strCheck_生产商 <> "" And .TextMatrix(intRow, mHeadCol.产地) <> "" Then
                    gstrSQL = "Select A.生产企业许可证, A.生产企业许可证效期,a.经营许可证, a.经营许可证效期, a.企业法人执照, a.企业法人执照效期 " & _
                              "From 材料生产商 A " & _
                              "Where A.名称 = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "校验生产商资质", .TextMatrix(intRow, mHeadCol.产地))
                    
                    strTmp_生产商 = ""
'                    strMsg_生产商 = ""
                    blnFlag = False
                    
                    If Not rsTmp.EOF Then
                        If zlStr.Nvl(rsTmp!生产企业许可证) = "" And InStr(strCheck_生产商 & ";", "生产企业许可证" & ";") > 0 Then
                            strTmp_生产商 = .TextMatrix(intRow, mHeadCol.产地) & "：" & "无生产企业许可证"
                            blnFlag = True
                        End If
                        If zlStr.Nvl(rsTmp!生产企业许可证效期) <> "" Then
                            If DateDiff("d", rsTmp!生产企业许可证效期, dateCurrent) > 0 And InStr(strCheck_生产商 & ";", "生产企业许可证效期" & ";") > 0 Then
                                strTmp_生产商 = IIf(strMsg_生产商 = "", .TextMatrix(intRow, mHeadCol.产地) & "：", strTmp_生产商 & ",") & "生产企业许可证已过期"
                                blnFlag = True
                            End If
                        End If
                        
                        If zlStr.Nvl(rsTmp!经营许可证) = "" And InStr(strCheck_生产商 & ";", "经营许可证" & ";") > 0 Then
                            strTmp_生产商 = .TextMatrix(intRow, mHeadCol.产地) & "：" & "无经营许可证"
                            blnFlag = True
                        End If
                        If zlStr.Nvl(rsTmp!经营许可证效期) <> "" Then
                            If DateDiff("d", rsTmp!经营许可证效期, dateCurrent) > 0 And InStr(strCheck_生产商 & ";", "经营许可证效期" & ";") > 0 Then
                                strTmp_生产商 = IIf(strMsg_生产商 = "", .TextMatrix(intRow, mHeadCol.产地) & "：", strTmp_生产商 & ",") & "经营许可证已过期"
                                blnFlag = True
                            End If
                        End If
                        
                        If zlStr.Nvl(rsTmp!企业法人执照) = "" And InStr(strCheck_生产商 & ";", "企业法人执照" & ";") > 0 Then
                            strTmp_生产商 = .TextMatrix(intRow, mHeadCol.产地) & "：" & "无企业法人执照"
                            blnFlag = True
                        End If
                        If zlStr.Nvl(rsTmp!企业法人执照效期) <> "" Then
                            If DateDiff("d", rsTmp!企业法人执照效期, dateCurrent) > 0 And InStr(strCheck_生产商 & ";", "企业法人执照效期" & ";") > 0 Then
                                strTmp_生产商 = IIf(strMsg_生产商 = "", .TextMatrix(intRow, mHeadCol.产地) & "：", strTmp_生产商 & ",") & "企业法人执照已过期"
                                blnFlag = True
                            End If
                        End If
                    End If
                    
                    If strTmp_生产商 <> "" Then
                        If InStr(1, str生产商列表, .TextMatrix(intRow, mHeadCol.产地)) = 0 Then
                            str生产商列表 = IIf(str生产商列表 = "", "", str生产商列表 & ",") & .TextMatrix(intRow, mHeadCol.产地)
                            
                            If intCount_生产商 <= 5 Then
                                strMsg_生产商 = strMsg_生产商 & strTmp_生产商 & vbCrLf
                            End If
                            intCount_生产商 = intCount_生产商 + 1
                        End If
                    End If
                    If blnFlag = True Then SetBilCheckFlag intRow, mHeadCol.产地, False
                    
                End If
                
                If strCheck_供应商 <> "" And .TextMatrix(intRow, mHeadCol.上次供应商) <> "" Then
                    gstrSQL = "Select 税务登记号, 许可证号, 执照号, 授权号, 质量认证号, 药监局备案号, 许可证效期, 执照效期, 授权期 " & _
                              "From 供应商 " & _
                              "Where (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And 名称 = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "供应商信息", .TextMatrix(intRow, mHeadCol.上次供应商))
                    
                    strTmp_供应商 = ""
'                    strMsg_供应商 = ""
                    blnFlag = False
                    
                    If Not rsTmp.EOF Then
                        If zlStr.Nvl(rsTmp!税务登记号) = "" And InStr(strCheck_供应商, "税务登记号") > 0 Then
                            strTmp_供应商 = .TextMatrix(intRow, mHeadCol.上次供应商) & "：" & "无税务登记号"
                            blnFlag = True
                        End If
                        
                        If zlStr.Nvl(rsTmp!许可证号) = "" And InStr(strCheck_供应商, "许可证号") > 0 Then
                            strTmp_供应商 = IIf(strTmp_供应商 = "", .TextMatrix(intRow, mHeadCol.上次供应商) & "：", strTmp_供应商 & ",") & "无许可证号"
                            blnFlag = True
                        End If
                        
                        If zlStr.Nvl(rsTmp!执照号) = "" And InStr(strCheck_供应商, "执照号") > 0 Then
                            strTmp_供应商 = IIf(strTmp_供应商 = "", .TextMatrix(intRow, mHeadCol.上次供应商) & "：", strTmp_供应商 & ",") & "无执照号"
                            blnFlag = True
                        End If
                        
                        If zlStr.Nvl(rsTmp!授权号) = "" And InStr(strCheck_供应商, "授权号") > 0 Then
                            strTmp_供应商 = IIf(strTmp_供应商 = "", .TextMatrix(intRow, mHeadCol.上次供应商) & "：", strTmp_供应商 & ",") & "无授权号"
                            blnFlag = True
                        End If
                        
                        If zlStr.Nvl(rsTmp!药监局备案号) = "" And InStr(strCheck_供应商, "药监局备案号") > 0 Then
                            strTmp_供应商 = IIf(strTmp_供应商 = "", .TextMatrix(intRow, mHeadCol.上次供应商) & "：", strTmp_供应商 & ",") & "无药监局备案号"
                            blnFlag = True
                        End If
                        
                        If zlStr.Nvl(rsTmp!许可证效期) <> "" Then
                            If DateDiff("d", rsTmp!许可证效期, dateCurrent) > 0 And InStr(strCheck_供应商, "许可证效期") > 0 Then
                                strTmp_供应商 = IIf(strTmp_供应商 = "", .TextMatrix(intRow, mHeadCol.上次供应商) & "：", strTmp_供应商 & ",") & "许可证已过期"
                                blnFlag = True
                            End If
                        End If
                        
                        If zlStr.Nvl(rsTmp!执照效期) <> "" Then
                            If DateDiff("d", rsTmp!执照效期, dateCurrent) > 0 And InStr(strCheck_供应商, "执照效期") > 0 Then
                                strTmp_供应商 = IIf(strTmp_供应商 = "", .TextMatrix(intRow, mHeadCol.上次供应商) & "：", strTmp_供应商 & ",") & "执照已过期"
                                blnFlag = True
                            End If
                        End If
                        
                        If zlStr.Nvl(rsTmp!授权期) <> "" Then
                            If DateDiff("d", rsTmp!执照效期, dateCurrent) > 0 And InStr(strCheck_供应商, "授权期") > 0 Then
                                strTmp_供应商 = IIf(strTmp_供应商 = "", .TextMatrix(intRow, mHeadCol.上次供应商) & "：", strTmp_供应商 & ",") & "授权已过期"
                                blnFlag = True
                            End If
                        End If
                    End If
                    
                    If strTmp_供应商 <> "" Then
                        If InStr(1, str供应商列表, .TextMatrix(intRow, mHeadCol.上次供应商)) = 0 Then
                            str供应商列表 = IIf(str供应商列表 = "", "", str供应商列表 & ",") & .TextMatrix(intRow, mHeadCol.上次供应商)
                            
                            If intCount_供应商 <= 5 Then
                                strMsg_供应商 = strMsg_供应商 & strTmp_供应商 & vbCrLf
                            End If
                            intCount_供应商 = intCount_供应商 + 1
                        End If
                    End If
                    If blnFlag = True Then SetBilCheckFlag intRow, mHeadCol.上次供应商, False
                End If
                 
                If strTmp_卫材 = "" And strTmp_生产商 = "" And strTmp_供应商 = "" Then
                    SetBilCheckFlag intRow, mHeadCol.校验, True
                End If
            End If
        Next
        .Redraw = True
    End With
    
    If strMsg_卫材 <> "" Then
        strMsg_卫材 = "卫材：" & vbCrLf & strMsg_卫材
        If intCount_卫材 > 5 Then strMsg_卫材 = strMsg_卫材 & vbCrLf & "....."
        strMsgInfo = IIf(strMsgInfo = "", "", strMsgInfo & vbCrLf) & strMsg_卫材
    End If
    
    If strMsg_生产商 <> "" Then
        strMsg_生产商 = "生产商：" & vbCrLf & strMsg_生产商
        If intCount_生产商 > 5 Then strMsg_生产商 = strMsg_生产商 & vbCrLf & "....."
        strMsgInfo = IIf(strMsgInfo = "", "", strMsgInfo & vbCrLf) & strMsg_生产商
    End If
    
    If strMsg_供应商 <> "" Then
        strMsg_供应商 = "供应商：" & vbCrLf & strMsg_供应商
        If intCount_供应商 > 5 Then strMsg_供应商 = strMsg_供应商 & vbCrLf & "....."
        strMsgInfo = IIf(strMsgInfo = "", "", strMsgInfo & vbCrLf) & strMsg_供应商
    End If
    
    If strMsgInfo <> "" Then
        strMsgInfo = "以下项目资质校验未通过，请检查：" & vbCrLf & strMsgInfo
        MsgBox strMsgInfo, vbOKOnly, gstrSysName
        CheckQualifications = False
        Exit Function
    End If
    
    CheckQualifications = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetBilCheckFlag(ByVal intRow As Integer, ByVal intCol As Integer, ByVal blnFlag As Boolean)
    '资质校验标记
    'blnFlag：True-在intRow行，intCol列打勾，表示所有项目都校验通过；False－不打勾，并且intRow行，intCol列上红色粗体标识
    Dim i As Integer
    With mshBill
        If blnFlag = False Then
            .Row = intRow
            i = .ColData(intCol)
            .ColData(intCol) = 0
            .Col = intCol
            .MsfObj.CellForeColor = vbRed
            .MsfObj.CellFontBold = True
            .ColData(intCol) = i
        Else
            .TextMatrix(intRow, intCol) = "√"
        End If
    End With
End Sub

Public Sub ShowCard(frmMain As Form, ByVal str单据号 As String, _
        ByVal int编辑状态 As Integer, Optional blnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1724)

    mblnSuccess = blnSuccess
    mblnChange = False
    mblnFirst = True

    Set mfrmMain = frmMain
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")

    If mint编辑状态 = 1 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If InStr(mstrPrivs, "单据打印") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If

    End If

    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号

End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

'查找
Private Sub cmdFind_Click()

    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindData mshBill, mHeadCol.材料, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindData mshBill, mHeadCol.材料, txtCode.Text, False
    ElseIf KeyCode = vbKeyEscape Then
        If Msf供应商选择.Visible Then
            Msf供应商选择.ZOrder 1
            Msf供应商选择.Visible = False
            Exit Sub
        End If
        Call CmdCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    
    If mint编辑状态 = 4 Then    '查看
        '打印
        Call FrmBillPrint.ShowMe(Me, glngSys, "zl1_bill_1724", 0, mintUnit, 1724, "卫材采购计划单", txtNO.Tag)
        '退出
        Unload Me
        Exit Sub
    End If

    If mint编辑状态 = 3 Then        '审核
        '资质校验
        If mblnFirstCheck = False Then
            mblnCheck = CheckQualifications
            mblnFirstCheck = True
            If mblnCheck = False Then
                Exit Sub
            End If
        End If
        
        If mblnCheck = False Then
            If mint校验方式 = 1 Then
                If MsgBox("部分卫材，生产商，供应商未通过校验，是否审核？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                
                If SaveCheckCard = False Then Exit Sub
            ElseIf mint校验方式 = 2 Then
                MsgBox "部分卫材，生产商，供应商未通过校验，不能审核！", vbOKOnly, gstrSysName
                Exit Sub
            End If
        End If
        
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
                '打印
                If InStr(mstrPrivs, "单据打印") <> 0 Then
                    ReportOpen gcnOracle, glngSys, "zl1_bill_1724", Me, "单据编号=" & txtNO.Tag, "单位=" & mintUnit, 2
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If

    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard

    If blnSuccess = True Then

        If IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '打印
            If InStr(mstrPrivs, "单据打印") <> 0 Then
                ReportOpen gcnOracle, glngSys, "zl1_bill_1724", Me, "单据编号=" & txtNO.Tag, "单位=" & mintUnit, 2
            End If
        End If
        If mint编辑状态 = 2 Then   '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    txt摘要.Text = ""
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "上一张单据的NO号：" & txtNO.Tag
End Sub

Private Function GetDeptRequestDataBill(ByVal strNOIn As String, ByVal lng库房ID As Long, _
    ByVal strStartDate As String, ByVal strEndDate As String, ByVal str分类IDIN As String) As Boolean
    '-----------------------------------------------------------------------------------------------------
    '功能:从部门申购中获取数据
    '参数:strNOIn-单据号
    '     库房ID-库房
    '    strStartDate-开始日期
    '    strEndDate-结束日期
    '    str分类IDIN-分类ID_IN
    '返回,设置成功,返回true,否则返回False:
    '-----------------------------------------------------------------------------------------------------

    Dim strSQL As String
    Dim rsplan As New Recordset
    Dim strSQL供应商 As String
    Dim rs供应商 As New Recordset
    Dim lngRecord As Long, lngProcess As Long
    Dim intLop As Integer
    strNOIn = Replace(strNOIn, "'", "")
    Me.MousePointer = vbHourglass
    mshBill.Redraw = False
    stbThis.Panels(2).Text = "正在计算"
    
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    Pic单据.Enabled = False

    err = 0: On Error GoTo ErrHand:
    If str分类IDIN <> "" Then
          strSQL = "" & _
                  "   Select  /*+ Rule*/ A.材料id,('['|| q.编码 || ']' || q.名称) as 材料信息,b.招标材料,q.规格,nvl(max(A.上次生产商),max(q.产地)) as 产地,nvl(Max(a.上次供应商),Max(g.名称)) as 上次供应商," & _
                  "      sum(nvl(A.请购数量,0)) as 请购数量," & _
                  "      sum(nvl(A.计划数量,0)) as 审批数量," & _
                          IIf(mintUnit = 0, "Q.计算单位", "B.包装单位") & " as 单位, " & _
                          IIf(mintUnit = 0, "1", "B.换算系数") & " as 换算系数, " & _
                  "       max(A.单价) as 单价," & _
                  "       sum(nvl(A.金额,0)) as 金额 " & _
                  "   From 材料计划内容 A,材料特性 B,材料采购计划 c,收费项目目录 Q,诊疗项目目录 M, " & _
                  "       Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) J, " & _
                  "       Table(Cast(f_Str2list([2]) As zlTools.t_StrList)) L, 供应商 G" & _
                  "   Where A.材料id=B.材料id and A.材料id=q.id And (q.站点=[6] or q.站点 is null) " & _
                  "         And a.计划id=c.id and c.单据=1" & IIf(lng库房ID = 0, "", " And C.库房id=[3]") & _
                  "         And (C.审核日期 between [4] and [5]) And C.No =L.Column_Value and Nvl(b.上次供应商id, 0) = g.Id(+)" & _
                  "         And B.诊疗id=M.id and M.分类id=J.Column_Value " & _
                  "   Group by A.材料id,q.编码 ,q.名称,q.规格,q.产地,b.招标材料,B.换算系数," & IIf(mintUnit = 0, "Q.计算单位", "B.包装单位")
      Else
          strSQL = "" & _
                  "   Select /*+ Rule*/ A.材料id, ('['|| q.编码 || ']' || q.名称) as 材料信息,B.招标材料,q.规格 ,nvl(max(A.上次生产商),q.产地) as 产地,nvl(Max(a.上次供应商),Max(g.名称)) as 上次供应商," & _
                  "      sum(nvl(A.请购数量,0)) as 请购数量," & _
                  "      sum(nvl(A.计划数量,0)) as 审批数量," & _
                         IIf(mintUnit = 0, "Q.计算单位", "B.包装单位") & " as 单位, " & _
                         IIf(mintUnit = 0, "1", "B.换算系数") & " as 换算系数, " & _
                  "      max(A.单价) as 单价," & _
                  "      sum(nvl(A.金额,0)) as 金额 " & _
                  "   From 材料计划内容 A,材料特性 B,材料采购计划 c,收费项目目录 Q," & _
                  "       Table(Cast(f_Str2list([2]) As zlTools.t_StrList)) L, 供应商 G" & _
                  "   Where A.材料id=B.材料id and A.材料id=q.id And (q.站点=[6] or q.站点 is null) " & _
                  "         And a.计划id=c.id and c.单据=1 " & IIf(lng库房ID = 0, "", " And C.库房id=[3]") & _
                  "         And (C.审核日期 between [4] and [5]) And C.No =L.Column_Value and Nvl(b.上次供应商id, 0) = g.Id(+)" & _
                  "   Group by A.材料id,q.编码 ,q.名称,q.规格,q.产地,b.招标材料,B.换算系数," & IIf(mintUnit = 0, "Q.计算单位", "B.包装单位") & _
                  "   "
      End If
    
    strSQL = "" & _
    "   Select  A.材料ID,A.材料信息,A.规格,A.招标材料,nvl(max(a.产地),a.产地) as 产地,A.上次供应商 as 上次供应商," & _
    "           A.请购数量,A.审批数量,A.单位,A.换算系数," & _
    "           nvl(max(b.上次采购价),a.单价)  as 单价," & _
    "           sum(nvl(B.实际数量,0)) as 库存数量" & _
    "   From (" & strSQL & ") A,药品库存 B " & _
    "   Where a.材料id=b.药品id(+) and b.性质(+)=1  " & IIf(lng库房ID = 0, "", " And B.库房id(+)=[3]") & _
    "   Group by A.材料ID,A.材料信息,A.规格,A.招标材料,a.产地,A.上次供应商,A.请购数量,A.审批数量,A.单位,A.换算系数,A.单价,A.金额 " & _
    "   order by 材料信息"
    
    Set rsplan = zlDatabase.OpenSQLRecord(strSQL, mstrCaption, str分类IDIN, strNOIn, lng库房ID, CDate(strStartDate), CDate(strEndDate & " 23:59:59"), gstrNodeNo)
    With rsplan
        lngRecord = .RecordCount
        If lngRecord = 0 Then
            mshBill.Redraw = True
            Me.MousePointer = vbDefault
            CmdSave.Enabled = True
            CmdCancel.Enabled = True
            Pic单据.Enabled = True
            Me.stbThis.Panels(2).Text = ""
            GetDeptRequestDataBill = True
            Exit Function
        End If
        lngProcess = 0
        If .RecordCount <> 0 Then .MoveFirst
        For intLop = 1 To .RecordCount
            mshBill.TextMatrix(intLop, 0) = Val(zlStr.Nvl(!材料ID))
            mshBill.TextMatrix(intLop, mHeadCol.材料) = zlStr.Nvl(!材料信息)
            mshBill.TextMatrix(intLop, mHeadCol.规格) = zlStr.Nvl(!规格)
            mshBill.TextMatrix(intLop, mHeadCol.产地) = zlStr.Nvl(!产地)
            mshBill.TextMatrix(intLop, mHeadCol.单位) = zlStr.Nvl(!单位)
            mshBill.TextMatrix(intLop, mHeadCol.前期数量) = ""
            mshBill.TextMatrix(intLop, mHeadCol.上期数量) = ""
            mshBill.TextMatrix(intLop, mHeadCol.库存数量) = Format(Val(zlStr.Nvl(!库存数量)) / Val(zlStr.Nvl(!换算系数)), mFMT.FM_数量)
            mshBill.TextMatrix(intLop, mHeadCol.计划数量) = Format(Val(zlStr.Nvl(!审批数量)) / Val(zlStr.Nvl(!换算系数)), mFMT.FM_数量)
            mshBill.TextMatrix(intLop, mHeadCol.单价) = Format(Val(zlStr.Nvl(!单价)) * Val(zlStr.Nvl(!换算系数)), mFMT.FM_成本价)
            mshBill.TextMatrix(intLop, mHeadCol.金额) = Format(Val(zlStr.Nvl(!单价)) * Val(zlStr.Nvl(!审批数量)), mFMT.FM_金额)
            mshBill.TextMatrix(intLop, mHeadCol.比例系数) = zlStr.Nvl(!换算系数)
            If zlStr.Nvl(!上次供应商) <> "" Then
                mshBill.TextMatrix(intLop, mHeadCol.上次供应商) = zlStr.Nvl(!上次供应商)
            Else
                strSQL供应商 = "Select b.上次供应商" & vbNewLine & _
                                        "From (Select Max(a.计划id) As 计划id From 材料计划内容 A Where a.材料id =[1] And a.上次供应商 Is Not Null) A, 材料计划内容 B, 供应商 C" & vbNewLine & _
                                        "Where a.计划id = b.计划id And b.上次供应商 = c.名称"

                Set rs供应商 = zlDatabase.OpenSQLRecord(strSQL供应商, "申购单取供应商", Val(zlStr.Nvl(!材料ID)))
                
                If rs供应商.RecordCount > 0 Then mshBill.TextMatrix(intLop, mHeadCol.上次供应商) = zlStr.Nvl(rs供应商!上次供应商)
            End If
            mshBill.TextMatrix(intLop, mHeadCol.中标材料) = zlStr.Nvl(!招标材料)
            
            Call Calc销量(Val(zlStr.Nvl(!材料ID)), intLop)
            
            If intLop >= mshBill.Rows - 1 Then mshBill.Rows = mshBill.Rows + 1
            lngProcess = lngProcess + 1
            Call ShowPercent(lngProcess / lngRecord)
            .MoveNext
        Next
        Call 显示合计金额
        .Close
    End With
    GetDeptRequestDataBill = True
    Call RefreshRowNO(mshBill, mHeadCol.序号, 1)
    Call 显示合计金额
    Me.MousePointer = vbDefault
    mshBill.Redraw = True
    CmdSave.Enabled = True
    Pic单据.Enabled = True
    CmdCancel.Enabled = True
    mshBill.Col = mHeadCol.计划数量
    Me.stbThis.Panels(2).Text = ""
    FS.StopFlash
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Activate()
    Dim intMonth As Integer

    If mblnFirst = False Then Exit Sub
    
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int简码方式 = Val(zlDatabase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram stbThis, gSystem_Para.int简码方式
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
    
    mblnFirst = False
    If mint编辑状态 = 1 Then
        Dim str分类ID As String, str剂型编码 As String
        Dim lng库房ID As Long, int计划类型 As Integer, int编制方法 As Integer
        
        If frmStuffPlanCondition.GetCondition(mfrmMain, str分类ID, lng库房ID, int计划类型, int编制方法, mbln下限, mint上限, mint下限, mstr供货商ID, mbln中标单位, mbln计划数量) = True Then
            mlng库房id = lng库房ID
            mint计划类型 = int计划类型
            mint编制方法 = int编制方法
            
            Select Case mint计划类型
                Case 1       '月计划
                    mstr期间 = Format(DateAdd("m", 1, sys.Currentdate), "yyyyMM")
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr期间, 1, 4) & "年" & Right(mstr期间, 2) & "月" & ") " & LblTitle.Tag & "采购计划"
'                    mshBill.TextMatrix(0, mHeadCol.上期销量) = "上月销量"
'                    mshBill.TextMatrix(0, mHeadCol.本期销量) = "本月销量"
                Case 2       '季计划
                    intMonth = Month(DateAdd("Q", 1, sys.Currentdate))
                    mstr期间 = Format(DateAdd("Q", 1, sys.Currentdate), "yyyy") & IIf(intMonth <= 3, 1, IIf(intMonth >= 10, 4, IIf(intMonth <= 9 And intMonth >= 7, 3, 2)))
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr期间, 1, 4) & "年" & Right(mstr期间, 1) & "季" & ")" & LblTitle.Tag & "采购计划"
'                    mshBill.TextMatrix(0, mHeadCol.上期销量) = "上季度销量"
'                    mshBill.TextMatrix(0, mHeadCol.本期销量) = "本季度销量"
                Case 3    '年计划
                    mstr期间 = Format(DateAdd("yyyy", 1, sys.Currentdate), "yyyy")
                    LblTitle.Caption = GetUnitName & "(" & mstr期间 & "年" & ")" & LblTitle.Tag & "采购计划"
'                    mshBill.TextMatrix(0, mHeadCol.上期销量) = "上年销量"
'                    mshBill.TextMatrix(0, mHeadCol.本期销量) = "本年销量"
                Case 4      '周计划
                    mstr期间 = Format(DateAdd("ww", 1, sys.Currentdate), "yyyyWW")
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr期间, 1, 4) & "年" & Right(mstr期间, 2) & "周" & ") " & LblTitle.Tag & "采购计划"
                    
            End Select
            If mint编制方法 = 5 Then
                Dim strStartDate As String, strEndDate As String, strNOIn As String
                '按部门申购编制采购计划
                 If FrmBillSelect.ShowCard(str分类ID, mlng库房id, mint计划类型, strNOIn, strStartDate, strEndDate) = False Then Unload Me: Exit Sub
                 If GetDeptRequestDataBill(strNOIn, mlng库房id, strStartDate, strEndDate, str分类ID) = False Then Exit Sub
                 Exit Sub
                 
            Else
                ReFreshALLStuff str分类ID, lng库房ID, int计划类型, int编制方法
            End If
        Else
            Unload Me
            Exit Sub
        End If
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    Else
'        mblnChange = False
        Select Case mintParallelRecord
            Case 1
                '正常
            Case 2
                '单据已被删除
                MsgBox "该单据已被删除，请检查！", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
            Case 3
                '修改的单据已被审核
                MsgBox "该单据已被其他人审核，请检查！", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
        End Select
    End If
    mblnStart = True
End Sub

Private Sub ReFreshALLStuff(ByVal str分类ID, _
    ByVal lng库房ID As Long, ByVal int计划类型 As Integer, ByVal int编制方法 As Integer)
        '---------------------------------------------------
        '--功能:对所有药品进行计划编制
        '--参数:
        '---------------------------------------------------
    Dim rsAllStuff As New ADODB.Recordset, rspurchase As New ADODB.Recordset
    Dim lngProcess  As Long, lngRecord As Long, lngRow As Long
    Dim blnOK As Boolean
    
    On Error GoTo ErrHandle
    Me.Refresh
    Me.MousePointer = vbHourglass
    mshBill.Redraw = False
    stbThis.Panels(2).Text = "正在计算"
    
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    Pic单据.Enabled = False

    Dim str单位 As String
    
    Select Case mintUnit
    Case 0
        str单位 = ",F.计算单位 单位,1 比例系数"
    Case Else
        str单位 = ",A.包装单位 单位,A.换算系数 比例系数"
    End Select

    '取指定条件的药品信息
    gstrSQL = "" & _
         " SELECT /*+ Rule*/ DISTINCT A.材料id 药品ID,A.招标材料,F.编码,NVL(B.名称,F.名称) AS 通用名称," & _
         "      F.规格" & str单位 & ",DECODE(A.成本价,NULL,NVL(A.指导批发价,0),0,NVL(A.指导批发价,0),NVL(A.成本价,0)) AS 单价,F.产地" & _
         " FROM 材料特性 A,收费项目别名 B,诊疗项目目录 C,诊疗分类目录 L,收费项目目录 F " & _
         IIf(str分类ID = "", "", ",Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) D ") & _
         " WHERE A.材料ID=F.ID And (f.站点=[3] or f.站点 is null) And A.诊疗ID=C.ID And C.分类ID=L.ID and L.类型 =7" & _
         "          And A.材料ID = B.收费细目ID(+) And B.性质(+)=3 " & _
         "          AND (F.撤档时间>=TO_DATE('3000-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS') OR F.撤档时间 IS NULL)" & _
                    IIf(str分类ID = "", IIf(mstr供货商ID <> "", "", " And L.ID Is NULL"), " AND L.ID =D.Column_Value ")

    '从药品库存中提取有库存的卫材的供应商与产地信息，无库存的卫材只取最后一次入库的供应商与产地信息
    
    If lng库房ID = 0 Then
        '如果是全库房，取所有库房库存，并从药品规格中取上次供应商和上次产地
        gstrSQL = "( " & gstrSQL & ") D," & _
                  " (Select a.药品id, c.Id As 上次供应商id, c.名称 As 供应商, b.上次产地, a.库存数量, a.平均售价 " & _
                " From (Select 药品id, Sum(实际数量) As 库存数量, " & _
                "              Decode(Sign(Sum(实际数量)), 1, Decode(Sign(Sum(实际金额)), 1, Sum(实际金额), 0) / Sum(实际数量), 0) 平均售价 " & _
                "       From 药品库存 " & _
                "       Where 性质 = 1 " & _
                "       Group By 药品id) A, 材料特性 B, (Select ID, 名称 From 供应商 Where Substr(类型, 5, 1) = 1) C " & _
                " Where a.药品id = b.材料id And b.上次供应商id = c.Id(+)) E "
    Else
        '取库存数量，及最大批次的供应商，上次产地
        gstrSQL = "( " & gstrSQL & ") D," & _
                  " (   Select A.药品ID,C.ID 上次供应商ID, C.名称 As 供应商, B.上次产地, A.库存数量, A.平均售价 " & _
                  "     From (  Select 库房id, 药品id, Sum(实际数量) As 库存数量, " & _
                  "                     Decode(Sign(Sum(实际数量)), 1, Decode(Sign(Sum(实际金额)), 1, Sum(实际金额), 0) / Sum(实际数量), 0) 平均售价 " & _
                  "             From 药品库存 " & _
                  "             Where 性质 = 1 " & IIf(lng库房ID = 0, "", " AND 库房ID= [1]") & _
                  "             Group By 库房id, 药品id) A, " & _
                  "          (  Select 库房id,药品id,批次,上次供应商ID,上次产地 From 药品库存 " & _
                  "             Where 性质 = 1 " & IIf(lng库房ID = 0, "", " AND 库房ID= [1]") & _
                  "                     And (药品ID,Nvl(批次, 0)) in  (Select 药品id,Nvl(Max(Nvl(批次, 0)), 0) 批次 From 药品库存 Where 性质 = 1 " & IIf(lng库房ID = 0, "", " AND 库房ID=[1] ") & " group by 药品id )" & _
                  "                                   ) B, " & _
                  "          (SELECT ID,名称 FROM 供应商 WHERE SUBSTR(类型,5,1)=1 ) C " & _
                  "     Where A.库房id = B.库房id And A.药品id = B.药品id And B.上次供应商id = C.ID(+) " & _
                  "     ) E "
    End If
    '加上提取材料储备限额的SQL
    gstrSQL = gstrSQL & _
            " ,  (  Select 材料id 药品ID,sum(nvl(上限,0)) 上限,sum(nvl(下限,0)) 下限 " & _
            "       From 材料储备限额  " & _
            "       " & IIf(lng库房ID = 0, "", " Where 库房ID=[1]") & _
            "       Group By 材料ID)     F"

    '联合所有（在最外层加上取材料储备限额.下限）
    gstrSQL = "" & _
        "   SELECT d.药品id,D.招标材料,e.上次供应商ID, d.编码, d.通用名称, d.规格, " & _
        "           DECODE (e.上次产地, NULL, d.产地, e.上次产地) AS 产地," & _
        "           d.单位,nvl(e.库存数量,0)/d.比例系数 as 库存数量 ,f.上限/d.比例系数 上限,f.下限/d.比例系数 下限 , d.单价 as 单价 , e.供应商,d.比例系数 from " & _
                gstrSQL & _
        " WHERE d.药品id = e.药品id (+) "
    '加上储备限额的判断，低于储备限额的药品才提取出来做采购计划
    gstrSQL = gstrSQL & " And d.药品ID=F.药品ID(+)"
    If mbln下限 Then
        '加上条件判断
        gstrSQL = "Select * From (" & gstrSQL & ") Where (库存数量<下限 and 下限<>0)"
    End If
    gstrSQL = gstrSQL & " Order by 编码"
        
    Set rsAllStuff = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng库房ID, str分类ID, gstrNodeNo)
    
    With rsAllStuff
        lngRecord = .RecordCount

        If lngRecord = 0 Then
            mshBill.Redraw = True
            Me.MousePointer = vbDefault
            CmdSave.Enabled = True
            CmdCancel.Enabled = True
            Pic单据.Enabled = True
            Me.stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        .MoveFirst
        Me.Refresh
        DoEvents
        Dim str上次供应商 As String
        
        lngRow = 0
        lngProcess = 1
        Do While Not .EOF
            blnOK = False
            str上次供应商 = ""
            If mstr供货商ID = "" Then
                blnOK = True
            Else
                If Val(zlStr.Nvl(!上次供应商id)) = 0 And mbln中标单位 Then
                    gstrSQL = "Select b.名称 from 材料中标单位 a,供应商 b,Table(cast(f_Num2List([2]) as zlTools.t_NumList)) C " & _
                              "Where a.材料ID=[1] and (b.站点=[2] or b.站点 is null) " & _
                              "    and a.单位id=b.id and a.单位ID=c.Column_Value "
                    Set rspurchase = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(zlStr.Nvl(!药品id)), mstr供货商ID, gstrNodeNo)
                    If rspurchase.RecordCount <> 0 Then
                        blnOK = True
                        str上次供应商 = zlStr.Nvl(rspurchase!名称)
                    End If
                Else
                     If "," & mstr供货商ID & "," Like "*," & Val(zlStr.Nvl(!上次供应商id)) & ",*" Then
                         blnOK = True
                     End If
                End If
            End If
            If blnOK Then
                    lngRow = lngRow + 1
                    mshBill.TextMatrix(lngRow, 0) = !药品id
                    mshBill.TextMatrix(lngRow, mHeadCol.材料) = "[" & !编码 & "]" & !通用名称
                    mshBill.TextMatrix(lngRow, mHeadCol.规格) = IIf(IsNull(!规格), "", !规格)
                    mshBill.TextMatrix(lngRow, mHeadCol.产地) = IIf(IsNull(!产地), "", !产地)
                    mshBill.TextMatrix(lngRow, mHeadCol.单位) = IIf(IsNull(!单位), "", !单位)
                    mshBill.TextMatrix(lngRow, mHeadCol.单价) = Format(Val(zlStr.Nvl(!单价)) * zlStr.Nvl(!比例系数, 1), mFMT.FM_成本价)
                    
                    mshBill.TextMatrix(lngRow, mHeadCol.上次供应商) = IIf(IsNull(!供应商), str上次供应商, !供应商)
                    mshBill.TextMatrix(lngRow, mHeadCol.库存数量) = Format(Val(zlStr.Nvl(!库存数量)), mFMT.FM_数量)
                    mshBill.TextMatrix(lngRow, mHeadCol.比例系数) = zlStr.Nvl(!比例系数, 1)
                    mshBill.TextMatrix(lngRow, mHeadCol.中标材料) = IIf(zlStr.Nvl(!招标材料) = 1, "√", "")
                    
                    mshBill.TextMatrix(lngRow, mHeadCol.存储上限) = Format(Val(zlStr.Nvl(!上限)), mFMT.FM_数量)
                    mshBill.TextMatrix(lngRow, mHeadCol.存储下限) = Format(Val(zlStr.Nvl(!下限)), mFMT.FM_数量)
                    
                    SetNumer !药品id, lng库房ID, Val(zlStr.Nvl(!库存数量)), lngRow, int计划类型, int编制方法
                    If lngRow = mshBill.Rows - 1 Then mshBill.Rows = mshBill.Rows + 1
            End If
            lngProcess = lngProcess + 1
            Call ShowPercent(lngProcess / lngRecord)
            .MoveNext
        Loop
    End With
    Call RefreshRowNO(mshBill, mHeadCol.序号, 1)
    Call 显示合计金额
    Me.MousePointer = vbDefault
    mshBill.Redraw = True
    CmdSave.Enabled = True
    Pic单据.Enabled = True
    CmdCancel.Enabled = True
    mshBill.Col = mHeadCol.计划数量
    Me.stbThis.Panels(2).Text = ""
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDate(ByVal int模式 As Integer, ByVal datCurrent As Date, _
        ByRef strBegin As String, ByRef strEnd As String) As Boolean
    Dim rsdate As New Recordset

    'int模式=1,月计划，2：季计划
    On Error GoTo ErrHandle
    GetDate = False
    If int模式 = 1 Then
        strBegin = Year(datCurrent) & "-" & String(2 - Len(Month(datCurrent)), "0") & Month(datCurrent) & "-01"
        gstrSQL = "select last_day(to_date([1],'yyyy-mm-dd')) from dual"
        Set rsdate = zlDatabase.OpenSQLRecord(gstrSQL, "GetDate", Format(datCurrent, "yyyy-mm-dd"))
        
        strEnd = Format(rsdate.Fields(0), "yyyy-mm-dd")
        rsdate.Close
    Else
        Select Case DatePart("Q", datCurrent)
            Case 1
                strBegin = Year(datCurrent) & "-01-01"
                strEnd = Year(datCurrent) & "-03-31"
            Case 2
                strBegin = Year(datCurrent) & "-04-01"
                strEnd = Year(datCurrent) & "-06-30"
            Case 3
                strBegin = Year(datCurrent) & "-07-01"
                strEnd = Year(datCurrent) & "-09-30"
            Case 4
                strBegin = Year(datCurrent) & "-10-01"
                strEnd = Year(datCurrent) & "-12-31"
        End Select
    End If
    GetDate = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'设置前期数量，上期数量，计划数量,金额等
Private Sub SetNumer(ByVal lng药品ID As Long, ByVal lng库房ID As Long, _
        ByVal num库存数量 As Double, ByVal intCurrentRow As Integer, _
        ByVal int计划类型 As Integer, ByVal int编制方法 As Integer)
    '---------------------------------------------------------------------------
    '--功能:确定耗用数量和计划数量
    '   1 )往年同期线性参照法：根据去前年同期药品的消耗情况，按线性规划原则预测消耗，对比库存产生采购计划供用户修改调整
    '   2 )临近期间平均参照法：以同年临近期间(前期、上期)的平均消耗预测消耗对比库存产生采购计划供用户修改调整；
    '   3 )药品储备参照法：根据药品储务下限与库存相减所得的差额为药品计划采购数；

    '--参数:
    '       int计划类型:1:月度计划,2.季度计划,3.年度计划,4.周度计划
    '       int编制方法:1 表示往年同期线性参照法,2 临近期间平均参照法,3.储备限额;4.日销售量
    '--返回:
    '---------------------------------------------------------------------------
    Dim num前期数量 As Double
    Dim num上期数量 As Double
    Dim num计划数量 As Double
    Dim num上限 As Double, num下限 As Double
    Dim lng天数 As Long

    Dim dat前期 As Date
    Dim dat上期 As Date
    Dim strBegin As String
    Dim strEnd As String
    Dim rsNum As New Recordset
    
    On Error GoTo ErrHandle
    With mshBill
        Select Case int编制方法
            Case 1      '往年同期线形参照   只有月度和季度计划
                dat前期 = DateAdd("m", Choose(int计划类型, 1, 3), DateAdd("yyyy", -2, sys.Currentdate))
                dat上期 = DateAdd("m", Choose(int计划类型, 1, 3), DateAdd("yyyy", -1, sys.Currentdate))
    
    
                If lng库房ID = 0 Then
                    GetDate int计划类型, dat前期, strBegin, strEnd
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 前期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "           and 单据 <>19 and 单据>=15 AND b.系数 = -1 " & _
                        "           AND 药品id+0 = [3]" & _
                        "           AND 日期 BETWEEN [1] and [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lng药品ID)
                    
                    If rsNum.EOF Then
                        num前期数量 = 0
                    Else
                        num前期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    GetDate int计划类型, dat上期, strBegin, strEnd
                    
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "           and 单据 <>19 and 单据>=15 AND b.系数 = -1 " & _
                        "           AND 药品id+0 = [3]" & _
                        "           AND 日期 BETWEEN [1] and [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lng药品ID)
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    GetDate int计划类型, dat前期, strBegin, strEnd
                    
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 前期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "           AND b.系数 = -1 " & _
                        "           and 库房id+0=[4]" & _
                        "           AND 药品id+0= [3] " & _
                        "           AND 日期 BETWEEN [1] and [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lng药品ID, lng库房ID)
    
                    If rsNum.EOF Then
                        num前期数量 = 0
                    Else
                        num前期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    
                    rsNum.Close
                    
                    GetDate int计划类型, dat上期, strBegin, strEnd
                    
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "       AND b.系数 = -1 " & _
                        "       and 库房id+0=[4]" & _
                        "       AND 药品id+0= [3]" & _
                        "       AND 日期 BETWEEN [1]  and  [2]"
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lng药品ID, lng库房ID)
    
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '把各单位转换成药库单位先
                num上期数量 = num上期数量 / .TextMatrix(intCurrentRow, mHeadCol.比例系数)
                num前期数量 = num前期数量 / .TextMatrix(intCurrentRow, mHeadCol.比例系数)
                '计划数量=2×上期数量－前期数量－库存数量
                If mbln计划数量 = True Then
                    num计划数量 = 2 * num上期数量 - num前期数量 - num库存数量
                    If num计划数量 < 0 Then num计划数量 = 0
                End If
    
                .TextMatrix(intCurrentRow, mHeadCol.前期数量) = Format(num前期数量, mFMT.FM_数量)
                .TextMatrix(intCurrentRow, mHeadCol.上期数量) = Format(num上期数量, mFMT.FM_数量)
                .TextMatrix(intCurrentRow, mHeadCol.计划数量) = IIf(Format(num计划数量, mFMT.FM_数量) = 0, "", Format(num计划数量, mFMT.FM_数量))
                .TextMatrix(intCurrentRow, mHeadCol.金额) = IIf(Format(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.单价) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.单价)), mFMT.FM_数量) = 0, "", Format(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.单价) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.单价)), mFMT.FM_数量))
            Case 2      '临近期间平均参照法
                dat前期 = Choose(int计划类型, DateAdd("m", -2, sys.Currentdate), DateAdd("m", -6, sys.Currentdate), DateAdd("yyyy", -2, sys.Currentdate), DateAdd("d", -14, sys.Currentdate))
                dat上期 = Choose(int计划类型, DateAdd("m", -1, sys.Currentdate), DateAdd("m", -3, sys.Currentdate), DateAdd("yyyy", -1, sys.Currentdate), DateAdd("d", -7, sys.Currentdate))
    
                If lng库房ID = 0 Then
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 前期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "           and 单据 <>19 and 单据>=15 AND b.系数 = -1 " & _
                        "           AND 药品id+0= [3]" & _
                        "           AND 日期 BETWEEN [1] and [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(Format(DateAdd(Choose(int计划类型, "m", "m", "m", "d"), Choose(int计划类型, -1, -3, -12, -7), dat前期), "yyyy-mm-dd hh:mm:ss")), CDate(Format(dat前期, "yyyy-mm-dd hh:mm:ss")), lng药品ID)
                    
                    If rsNum.EOF Then
                        num前期数量 = 0
                    Else
                        num前期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "       and 单据 <>19 and 单据>=15 AND b.系数 = -1 " & _
                        "       AND 药品id+0= [3]" & _
                        "       AND 日期 BETWEEN [1] " & _
                        "       and [2]"
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(Format(DateAdd(Choose(int计划类型, "m", "m", "m", "d"), Choose(int计划类型, -1, -3, -12, -7), dat上期), "yyyy-mm-dd hh:mm:ss")), CDate(Format(dat上期, "yyyy-mm-dd hh:mm:ss")), lng药品ID)
    
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 前期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "       AND b.系数 = -1 " & _
                        "       and a.库房id+0=[4]" & _
                        "       AND 药品id+0=[3] " & _
                        "       AND 日期 BETWEEN [1] and  [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(Format(DateAdd(Choose(int计划类型, "m", "m", "m", "d"), Choose(int计划类型, -1, -3, -12, -7), dat前期), "yyyy-mm-dd hh:mm;ss")), CDate(Format(dat前期, "yyyy-mm-dd hh:mm:ss")), lng药品ID, lng库房ID)
    
                    If rsNum.EOF Then
                        num前期数量 = 0
                    Else
                        num前期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "           AND b.系数 = -1 " & _
                        "           and a.库房id+0=[4]" & _
                        "           AND 药品id+0=[3] " & _
                        "           AND 日期 BETWEEN [1]  and [2]"
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(Format(DateAdd(Choose(int计划类型, "m", "m", "m", "d"), Choose(int计划类型, -1, -3, -12, -7), dat上期), "yyyy-mm-dd hh:mm;ss")), CDate(Format(dat上期, "yyyy-mm-dd hh:mm:ss")), lng药品ID, lng库房ID)
    
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '把各单位转换成药库单位先
                num上期数量 = num上期数量 / .TextMatrix(intCurrentRow, mHeadCol.比例系数)
                num前期数量 = num前期数量 / .TextMatrix(intCurrentRow, mHeadCol.比例系数)
                
                '计划数量 = (前期数量 + 上期数量) / 2 - 库存数量
                If mbln计划数量 = True Then
                    num计划数量 = (num上期数量 + num前期数量) / 2 - num库存数量
                    If num计划数量 < 0 Then num计划数量 = 0
                End If
                .TextMatrix(intCurrentRow, mHeadCol.前期数量) = Format(num前期数量, mFMT.FM_数量)
                .TextMatrix(intCurrentRow, mHeadCol.上期数量) = Format(num上期数量, mFMT.FM_数量)
                .TextMatrix(intCurrentRow, mHeadCol.计划数量) = IIf(Format(num计划数量, mFMT.FM_数量) = 0, "", Format(num计划数量, mFMT.FM_数量))
                .TextMatrix(intCurrentRow, mHeadCol.金额) = IIf(Format(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.单价) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.单价)), mFMT.FM_数量) = 0, "", Format(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.单价) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.单价)), mFMT.FM_数量))
    
            Case 3      '药品储备定额参照法
                If lng库房ID = 0 Then
                    gstrSQL = "select sum(上限) as  上限 from 材料储备限额  where 材料id=[1]"
                Else
                    gstrSQL = "select 上限 from 材料储备限额  where 材料id=[1] and 库房id=[2]"
    
                End If
                Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng药品ID, lng库房ID)
    
                If rsNum.EOF Then
                    num上限 = 0
                Else
                    num上限 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                End If
    
                '把各单位转换成药库单位先
                num上限 = num上限 / .TextMatrix(intCurrentRow, mHeadCol.比例系数)
                '计划数量=储备下限－库存数量
                If mbln计划数量 = True Then
                    num计划数量 = IIf(num上限 > num库存数量, num上限 - num库存数量, 0)
                End If
                .TextMatrix(intCurrentRow, mHeadCol.计划数量) = IIf(Format(num计划数量, mFMT.FM_数量) = 0, "", Format(num计划数量, mFMT.FM_数量))
                .TextMatrix(intCurrentRow, mHeadCol.金额) = IIf(Format(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.单价) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.单价)), mFMT.FM_数量) = 0, "", Format(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.单价) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.单价)), mFMT.FM_数量))
            Case 4  '日销售量
                dat前期 = Choose(int计划类型, DateAdd("m", -2, sys.Currentdate), DateAdd("m", -6, sys.Currentdate), DateAdd("yyyy", -2, sys.Currentdate), DateAdd("d", -14, sys.Currentdate))
                dat上期 = Choose(int计划类型, DateAdd("m", -1, sys.Currentdate), DateAdd("m", -3, sys.Currentdate), DateAdd("yyyy", -1, sys.Currentdate), DateAdd("d", -7, sys.Currentdate))
                GetDate int计划类型, dat上期, strBegin, strEnd
                lng天数 = CDate(Format(strEnd, "yyyy-MM-DD")) - CDate(Format(strBegin, "yyyy-MM-DD")) + 1
                If lng天数 <= 0 Then lng天数 = 1
                
                If lng库房ID = 0 Then
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "           and 单据 <>19 and 单据>=15 AND b.系数 = -1 " & _
                        "           AND 药品id+0 =[3] " & _
                        "           AND 日期 BETWEEN [1]   and  [2] "
                
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lng药品ID)
                            
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " & _
                        "   FROM 药品收发汇总 a, 药品入出类别 b " & _
                        "   Where a.类别id = b.id " & _
                        "       AND b.系数 = -1 " & _
                        "       and 库房id+0=[4]" & _
                        "       AND 药品id+0=[3] " & _
                        "       AND 日期 BETWEEN [1] and [2]"
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lng药品ID, lng库房ID)
    
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '把各单位转换成药库单位先
                num上期数量 = num上期数量 / .TextMatrix(intCurrentRow, mHeadCol.比例系数)
                num上限 = num上期数量 / lng天数 * mint上限
                num下限 = num上期数量 / lng天数 * mint下限
                '计划数量=2×上期数量－前期数量－库存数量
                If mbln计划数量 = True Then
                    If num库存数量 < num下限 Then
                        num计划数量 = num上限 - num库存数量
                    Else
                        num计划数量 = 0
                    End If
                    If num计划数量 < 0 Then num计划数量 = 0
                End If
    
                .TextMatrix(intCurrentRow, mHeadCol.前期数量) = Format(num前期数量, mFMT.FM_数量)
                .TextMatrix(intCurrentRow, mHeadCol.上期数量) = Format(num上期数量, mFMT.FM_数量)
                .TextMatrix(intCurrentRow, mHeadCol.计划数量) = IIf(Format(num计划数量, mFMT.FM_数量) = 0, "", Format(num计划数量, mFMT.FM_数量))
                .TextMatrix(intCurrentRow, mHeadCol.金额) = IIf(Format(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.单价) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.单价)), mFMT.FM_数量) = 0, "", Format(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.单价) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.单价)), mFMT.FM_数量))
        End Select
    
        Call Calc销量(lng药品ID, intCurrentRow)
        
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Calc销量(ByVal lng药品ID As Long, ByVal intCurrentRow As Integer)
    '分别计算上期和本期的销售量
    '取上期的区间范围
    Dim strBegin As String
    Dim strEnd As String
    Dim rsNum As ADODB.Recordset
    
    On Error GoTo ErrHandle
    With mshBill
        Select Case mint计划类型
            '1:月度计划,2.季度计划,3.年度计划,4.周度计划
            Case 1
                '上月时间范围
                strBegin = Format(DateAdd("m", -1, CDate(mstrNow)), "YYYY-MM") & "-01"
                strEnd = Format(DateAdd("d", -1, CDate(Format(CDate(mstrNow), "YYYY-MM") & "-01")), "YYYY-MM-DD") & " 23:59:59"
            Case 2
                '上季度时间范围
                Select Case DatePart("Q", CDate(mstrNow))
                    Case 1
                        strBegin = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-10-01"
                        strEnd = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-12-31 23:59:59"
                    Case 2
                        strBegin = Format(mstrNow, "YYYY") & "-01-01"
                        strEnd = Format(mstrNow, "YYYY") & "-03-31 23:59:59"
                     Case 3
                        strBegin = Format(mstrNow, "YYYY") & "-04-01"
                        strEnd = Format(mstrNow, "YYYY") & "-06-30 23:59:59"
                    Case 4
                        strBegin = Format(mstrNow, "YYYY") & "-07-01"
                        strEnd = Format(mstrNow, "YYYY") & "-09-30 23:59:59"
                End Select
            Case 3
                '上年度时间范围
                strBegin = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-01-01"
                strEnd = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-12-31 23:59:59"
            Case 4
                '上周时间范围
                strBegin = Format(DateAdd("d", 2 - Weekday(CDate(mstrNow)) - 7, CDate(mstrNow)), "YYYY-mm-dd")
                strEnd = Format(DateAdd("d", 8 - Weekday(CDate(mstrNow)) - 7, CDate(mstrNow)), "YYYY-mm-dd") & " 23:59:59"
                
        End Select
            
        '计算上期销售量（不要求精确值，用药品收发汇总统计）
        gstrSQL = "Select -Sum(Nvl(数量, 0)) As 销售数量 " & _
            " From 药品收发汇总" & _
            " Where 类别id + 0 In (19,20,21) And 药品id+0=[1] And 日期 Between [2] And [3] "
        Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng药品ID, CDate(strBegin), CDate(strEnd))
        If rsNum.RecordCount > 0 Then
            .TextMatrix(intCurrentRow, mHeadCol.上期销量) = Format(Nvl(rsNum!销售数量, 0) / Val(.TextMatrix(intCurrentRow, mHeadCol.比例系数)), mFMT.FM_数量)
        End If
        
        '取本期的区间范围
        Select Case mint计划类型
            '1:月度计划,2.季度计划,3.年度计划
            Case 1
                '本月时间范围
                strBegin = Format(mstrNow, "YYYY-MM") & "-01"
            Case 2
                '本季度时间范围
                Select Case DatePart("Q", CDate(mstrNow))
                    Case 1
                        strBegin = Format(mstrNow, "YYYY") & "-01-01"
                    Case 2
                        strBegin = Format(mstrNow, "YYYY") & "-04-01"
                    Case 3
                        strBegin = Format(mstrNow, "YYYY") & "-07-01"
                    Case 4
                        strBegin = Format(mstrNow, "YYYY") & "-10-01"
                End Select
            Case 3
                '本年度时间范围
                strBegin = Format(mstrNow, "YYYY") & "-01-01"
            Case 4
                '本周时间范围
                strBegin = Format(DateAdd("d", 2 - Weekday(CDate(mstrNow)), CDate(mstrNow)), "YYYY-mm-dd")
        End Select
        
        '本期结束时间截止到当日
        strEnd = Format(mstrNow, "YYYY-MM-DD") & " 23:59:59"
            
        '计算本期销售量（不要求精确值，用药品收发汇总统计）
        gstrSQL = "Select -Sum(Nvl(数量, 0)) As 销售数量 " & _
            " From 药品收发汇总" & _
            " Where 类别id + 0 In (19,20,21) And 药品id+0=[1] And 日期 Between [2] And [3] "
        Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng药品ID, CDate(strBegin), CDate(strEnd))
        If rsNum.RecordCount > 0 Then
            .TextMatrix(intCurrentRow, mHeadCol.本期销量) = Format(zlStr.Nvl(rsNum!销售数量, 0) / Val(.TextMatrix(intCurrentRow, mHeadCol.比例系数)), mFMT.FM_数量)
        End If
    End With
        
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowPercent(sngPercent As Single)
'功能:在状态条上根据百分比显示当前处理进度(█)
    Dim intAll As Integer
    intAll = stbThis.Panels(2).Width / TextWidth("█") - 4
    stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
End Sub

Private Sub Form_Load()
    Dim strReg As String
    mFMT.FM_金额 = GetDigit

    mintUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mstrNow = Format(sys.Currentdate, "yyyy-mm-dd")
    
    mint校验方式 = Val(Mid(zlDatabase.GetPara("资质校验", glngSys, mlngModule, ""), 1, 1))
    
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
     
'    mintUnit = GetUnit()
    txtNO = mstr单据号
    txtNO.Tag = txtNO
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Call initCard
    
    RestoreWinState Me, App.ProductName, mstrCaption
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshBill
        .ColWidth(mHeadCol.单价) = IIf(mblnCostView = True, 1000, 0)
        .ColWidth(mHeadCol.金额) = IIf(mblnCostView = True, 1000, 0)
    End With
    
    mshBill.ColWidth(mHeadCol.校验) = IIf(mint编辑状态 = 3 And mint校验方式 = 1, 500, 0)
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim intRecordCount As Integer
    Dim str单位 As String
    Dim strOrder As String, strCompare As String
    
    On Error GoTo ErrHandle
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strCompare = Mid(strOrder, 1, 1)

    '库房
    Select Case mint编辑状态
        Case 1
            Txt填制人 = gstrUserName
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
            strUnit = "包装单位"
            Select Case mintUnit
            Case 0
                str单位 = ",j.计算单位 单位,1 比例系数"
            Case Else
                str单位 = ",m.包装单位 单位,m.换算系数 比例系数"
            End Select
            
            initGrid
            
            gstrSQL = "Select 库房id From  材料采购计划 where nvl(单据,0)=0 and NO=[1] and rownum=1 "
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            mlng库房id = Val(zlStr.Nvl(rsInitCard!库房id))
            
            gstrSQL = "" & _
                "   SELECT a.id,nvl(a.库房id,0) as 库房id,nvl(c.名称,'全院') AS 库房,a.no, a.计划类型,a.期间, a.编制方法, a.编制人," & _
                "           TO_CHAR (a.编制日期, 'yyyy-mm-dd HH24:MI:SS') AS 编制日期, a.审核人," & _
                "           TO_CHAR (a.审核日期, 'yyyy-mm-dd HH24:MI:SS') AS 审核日期,a.编制说明," & _
                "           b.序号,b.材料id 药品id,m.招标材料 ,F.上限,F.下限,J.编码,J.名称 通用名称, J.规格" & str单位 & ", b.前期数量, b.上期数量,b.上期销量,b.本期销量, b.库存数量, b.计划数量, b.单价, b.金额, b.上次供应商,b.上次生产商 " & _
                "   FROM 材料采购计划 a, 材料计划内容 b,部门表 c,材料特性 M,收费项目目录 J," & _
                "       ( Select 材料id ,sum(nvl(上限,0)) 上限,sum(nvl(下限,0)) 下限 " & _
                "               From 材料储备限额  " & _
                "               " & IIf(mlng库房id = 0, "", " Where 库房ID=[2]") & _
                "               Group By 材料ID ) F " & _
                "   Where a.id = b.计划id and b.材料ID=f.材料id(+) and nvl(a.库房id,0)=c.id(+) " & _
                "          and b.材料id=m.材料id and m.材料id=J.id And (j.站点=[3] or j.站点 is null) And nvl(a.单据,0)=0 AND a.no = [1]" & _
                "   Order by " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "编码", "通用名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号, mlng库房id, gstrNodeNo)
                '"       (   SELECT DISTINCT a.材料id as 药品id,c.编码,C.名称  AS 通用名称,c.规格,c.计算单位 as 散装单位,A.包装单位,a.换算系数 " & _
                "           FROM 材料特性 a, 收费项目别名 b, 收费项目目录 c " & _
                "           WHERE a.材料id = b.收费细目ID(+) and B.性质(+)=3  AND a.材料id = c.ID" & _
                "        ) d
                
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If

            intRecordCount = rsInitCard.RecordCount

            Txt填制人 = rsInitCard!编制人
            If mint编辑状态 = 2 Then
                Txt填制人 = gstrUserName
            End If
            Txt填制日期 = Format(rsInitCard!编制日期, "yyyy-mm-dd hh:mm:ss")

            Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
            Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            txt摘要.Text = IIf(IsNull(rsInitCard!编制说明), "", rsInitCard!编制说明)
            txt计划类型 = Choose(rsInitCard!计划类型, "月度计划", "季度计划", "年度计划", "周度计划")
            txt编制方法 = Choose(rsInitCard!编制方法, "往年同期线形参照法", "临近期间平均参照法", "材料储备定额参照法", "卫材日销售量参照法", "部门申购参照法")
            mint计划类型 = rsInitCard!计划类型
            mint编制方法 = rsInitCard!编制方法
            mlng库房id = rsInitCard!库房id
            mlng计划ID = rsInitCard!Id

            mstr期间 = rsInitCard!期间
            Select Case mint计划类型
                Case 1       '月计划
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr期间, 1, 4) & "年" & Right(mstr期间, 2) & "月" & ") " & rsInitCard!库房 & "采购计划"
                Case 2       '季计划
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr期间, 1, 4) & "年" & Right(mstr期间, 1) & "季" & ")" & rsInitCard!库房 & "采购计划"
                Case 3       '年计划
                    LblTitle.Caption = GetUnitName & "(" & mstr期间 & "年" & ")" & rsInitCard!库房 & "采购计划"
                Case 4       '周计划
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr期间, 1, 4) & "年" & Right(mstr期间, 2) & "周" & ") " & rsInitCard!库房 & "采购计划"
            End Select

            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If

            With mshBill
'                Select Case mint计划类型
'                    Case 1
'                        .TextMatrix(0, mHeadCol.上期销量) = "上月销量"
'                        .TextMatrix(0, mHeadCol.本期销量) = "本月销量"
'                    Case 2
'                        .TextMatrix(0, mHeadCol.上期销量) = "上季度销量"
'                        .TextMatrix(0, mHeadCol.本期销量) = "本季度销量"
'                    Case Else
'                        .TextMatrix(0, mHeadCol.上期销量) = "上年销量"
'                        .TextMatrix(0, mHeadCol.本期销量) = "本年销量"
'                End Select
                For intRow = 1 To intRecordCount

                    .TextMatrix(intRow, 0) = rsInitCard!药品id
                    .TextMatrix(intRow, mHeadCol.材料) = "[" & rsInitCard!编码 & "]" & rsInitCard!通用名称
                    .TextMatrix(intRow, mHeadCol.规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mHeadCol.上次供应商) = IIf(IsNull(rsInitCard!上次供应商), "", rsInitCard!上次供应商)
                    .TextMatrix(intRow, mHeadCol.产地) = IIf(IsNull(rsInitCard!上次生产商), "", rsInitCard!上次生产商)
                    .TextMatrix(intRow, mHeadCol.单位) = rsInitCard!单位
                    .TextMatrix(intRow, mHeadCol.比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mHeadCol.中标材料) = IIf(Val(zlStr.Nvl(rsInitCard!招标材料)) = 1, "√", "")
                    .TextMatrix(intRow, mHeadCol.存储上限) = Format(Val(zlStr.Nvl(rsInitCard!上限)) / rsInitCard!比例系数, mFMT.FM_数量)
                    .TextMatrix(intRow, mHeadCol.存储下限) = Format(Val(zlStr.Nvl(rsInitCard!下限)) / rsInitCard!比例系数, mFMT.FM_数量)
                    .TextMatrix(intRow, mHeadCol.前期数量) = Format(Val(zlStr.Nvl(rsInitCard!前期数量)) / rsInitCard!比例系数, mFMT.FM_数量)
                    .TextMatrix(intRow, mHeadCol.上期数量) = Format(Val(zlStr.Nvl(rsInitCard!上期数量)) / rsInitCard!比例系数, mFMT.FM_数量)
                    .TextMatrix(intRow, mHeadCol.库存数量) = Format(Val(zlStr.Nvl(rsInitCard!库存数量)) / rsInitCard!比例系数, mFMT.FM_数量)
                    
                    .TextMatrix(intRow, mHeadCol.上期销量) = Format(Val(zlStr.Nvl(rsInitCard!上期销量)) / rsInitCard!比例系数, mFMT.FM_数量)
                    .TextMatrix(intRow, mHeadCol.本期销量) = Format(Val(zlStr.Nvl(rsInitCard!本期销量)) / rsInitCard!比例系数, mFMT.FM_数量)
                    
                    .TextMatrix(intRow, mHeadCol.计划数量) = IIf(Format(Val(zlStr.Nvl(rsInitCard!计划数量)), mFMT.FM_数量) = 0, "", Format(rsInitCard!计划数量 / rsInitCard!比例系数, mFMT.FM_数量))
                    .TextMatrix(intRow, mHeadCol.单价) = Format(Val(zlStr.Nvl(rsInitCard!单价)) * rsInitCard!比例系数, mFMT.FM_成本价)
                    .TextMatrix(intRow, mHeadCol.金额) = IIf(Format(Val(zlStr.Nvl(rsInitCard!金额)), mFMT.FM_金额) = 0, "", Format(Val(zlStr.Nvl(rsInitCard!金额)), mFMT.FM_金额))
                    If intRow = .Rows - 1 Then .Rows = .Rows + 1
                    rsInitCard.MoveNext
                Next
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mHeadCol.序号, 1)
    Call 显示合计金额
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'初始化编辑控件
Private Sub initGrid()
    Dim intCol As Integer

    With mshBill
        .Active = True
        .Cols = mconIntColS
        .MsfObj.FixedCols = 2

        .TextMatrix(0, mHeadCol.序号) = "序号"
        .TextMatrix(0, mHeadCol.校验) = "校验"
        .TextMatrix(0, mHeadCol.材料) = "材料名称与编码"
        .TextMatrix(0, mHeadCol.规格) = "规格"
        .TextMatrix(0, mHeadCol.产地) = "产地"
        .TextMatrix(0, mHeadCol.单位) = "单位"
        .TextMatrix(0, mHeadCol.比例系数) = "比例系数"
        .TextMatrix(0, mHeadCol.中标材料) = "中标材料"
        .TextMatrix(0, mHeadCol.存储上限) = "存储上限"
        .TextMatrix(0, mHeadCol.存储下限) = "存储下限"

        .TextMatrix(0, mHeadCol.前期数量) = "前期数量"
        .TextMatrix(0, mHeadCol.上期数量) = "上期数量"
        .TextMatrix(0, mHeadCol.库存数量) = "库存数量"
        .TextMatrix(0, mHeadCol.上期销量) = "上期销量"
        .TextMatrix(0, mHeadCol.本期销量) = "本期销量"
        
        .TextMatrix(0, mHeadCol.计划数量) = "计划数量"
        .TextMatrix(0, mHeadCol.单价) = "成本价"
        .TextMatrix(0, mHeadCol.金额) = "成本金额"
        .TextMatrix(0, mHeadCol.上次供应商) = "供应商"

        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mHeadCol.序号) = "1"

        .ColWidth(mHeadCol.序号) = 500
        .ColWidth(mHeadCol.校验) = IIf(mint编辑状态 = 3 And mint校验方式 = 1, 500, 0)
        .ColWidth(mHeadCol.材料) = 2000
        .ColWidth(mHeadCol.规格) = 900
        .ColWidth(mHeadCol.产地) = 800
        .ColWidth(mHeadCol.单位) = 500
        .ColWidth(mHeadCol.前期数量) = 1100
        .ColWidth(mHeadCol.上期数量) = 1100
        .ColWidth(mHeadCol.库存数量) = 1100
        .ColWidth(mHeadCol.上期销量) = 1100
        .ColWidth(mHeadCol.本期销量) = 1100
        .ColWidth(mHeadCol.计划数量) = 1100
        .ColWidth(mHeadCol.中标材料) = 800
        .ColWidth(mHeadCol.存储上限) = 1000
        .ColWidth(mHeadCol.存储下限) = 1000
        
        .ColWidth(mHeadCol.单价) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(mHeadCol.金额) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mHeadCol.上次供应商) = 900
        .ColWidth(mHeadCol.比例系数) = 0
        .ColWidth(0) = 0

        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择
        For intCol = 0 To .Cols - 1
            .ColData(intCol) = 5
        Next

        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txt摘要.Enabled = True
            .ColData(mHeadCol.材料) = 1
            .ColData(mHeadCol.计划数量) = 4
            .ColData(mHeadCol.单价) = 4

            .ColData(mHeadCol.产地) = 1
            .ColData(mHeadCol.上次供应商) = 1
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 4 Then
            txt摘要.Enabled = False
            .ColData(mHeadCol.计划数量) = 0
        End If
        
        .ColData(mHeadCol.校验) = IIf(mint编辑状态 = 3 And mint校验方式 = 1, 0, 5)
        
        .ColAlignment(mHeadCol.校验) = flexAlignCenterCenter
        .ColAlignment(mHeadCol.材料) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.规格) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.产地) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.单位) = flexAlignCenterCenter
        .ColAlignment(mHeadCol.前期数量) = flexAlignRightCenter
        .ColAlignment(mHeadCol.上期数量) = flexAlignRightCenter
        .ColAlignment(mHeadCol.库存数量) = flexAlignRightCenter
        .ColAlignment(mHeadCol.上期销量) = flexAlignRightCenter
        .ColAlignment(mHeadCol.本期销量) = flexAlignRightCenter
        .ColAlignment(mHeadCol.计划数量) = flexAlignRightCenter
        .ColAlignment(mHeadCol.单价) = flexAlignRightCenter
        .ColAlignment(mHeadCol.金额) = flexAlignRightCenter
        .ColAlignment(mHeadCol.上次供应商) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.中标材料) = 4
        .ColAlignment(mHeadCol.存储上限) = 7
        .ColAlignment(mHeadCol.存储下限) = 7

        .PrimaryCol = mHeadCol.材料
        .LocateCol = mHeadCol.材料
        If InStr(1, "34", mint编辑状态) <> 0 Then .ColData(mHeadCol.材料) = 0
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub

    With Pic单据
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With

    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic单据.Width
    End With


    With mshBill
        .Left = 200
        .Width = Pic单据.Width - .Left * 2
    End With
    With txtNO
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With

    txt编制方法.Left = mshBill.Left + mshBill.Width - txt编制方法.Width
    lbl编制方法.Left = txt编制方法.Left - lbl编制方法.Width - 100


    Lbl计划类型.Left = mshBill.Left

    txt计划类型.Left = Lbl计划类型.Left + Lbl计划类型.Width + 100

    With Lbl填制人
        .Top = Pic单据.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With

    With Txt填制人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl填制人.Left + Lbl填制人.Width + 100
    End With

    With Lbl填制日期
        .Top = Lbl填制人.Top
        .Left = Txt填制人.Left + Txt填制人.Width + 250
    End With

    With Txt填制日期
        .Top = Lbl填制日期.Top - 80
        .Left = Lbl填制日期.Left + Lbl填制日期.Width + 100
    End With

    With Txt审核日期
        .Top = Lbl填制人.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With

    With Lbl审核日期
        .Top = Lbl填制人.Top
        .Left = Txt审核日期.Left - 100 - .Width
    End With

    With Txt审核人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl审核日期.Left - 200 - .Width
    End With

    With Lbl审核人
        .Top = Lbl填制人.Top
        .Left = Txt审核人.Left - 100 - .Width
    End With

    With txt摘要
        .Top = Lbl填制人.Top - 140 - .Height
        .Left = Txt填制人.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With

    With lbl摘要
        .Top = txt摘要.Top + 50
        .Left = txt摘要.Left - .Width - 100
    End With

    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
        .Width = mshBill.Width
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If

    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With

    With CmdCancel
        .Left = Pic单据.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic单据.Top + Pic单据.Height + 100
    End With

    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With

    With cmdHelp
        .Left = Pic单据.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With

    With cmdFind
        .Top = CmdCancel.Top
    End With

    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mblnFirstCheck = False
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
    mblnStart = False
End Sub

Private Function SaveCheck() As Boolean
    Dim str审核人 As String

    mblnSave = False
    SaveCheck = False

    str审核人 = gstrUserName

    On Error GoTo ErrHandle
    'zl_材料计划管理_VERIFY( /*ID_IN*/, /*审核人_IN*/ );
    gstrSQL = "zl_材料计划管理_VERIFY('" & mlng计划ID & "','" & str审核人 & "')"
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption

    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    'MsgBox "审核失败！", vbInformation, gstrSysName
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Sub Msf供应商选择_DblClick()
    Dim blnCancel As Boolean
    With mshBill
        .Text = Msf供应商选择.TextMatrix(Msf供应商选择.Row, 2)
        .TextMatrix(.Row, mHeadCol.上次供应商) = Msf供应商选择.TextMatrix(Msf供应商选择.Row, 2)
    End With
    Msf供应商选择.Visible = False
    mshBill.SetFocus
    Call SendKeys("{ENTER}")
End Sub

Private Sub Msf供应商选择_GotFocus()
    If Msf供应商选择.Rows - 1 = 1 Then Call Msf供应商选择_DblClick
End Sub

Private Sub Msf供应商选择_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Msf供应商选择_DblClick
    End If
End Sub

Private Sub Msf供应商选择_LostFocus()
    Msf供应商选择.ZOrder 1
    Msf供应商选择.Visible = False
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mHeadCol.序号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mHeadCol.序号, mshBill.Row)
    Call 显示合计金额
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mHeadCol.材料) = 0 Then
        Exit Sub
    End If
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint编辑状态) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("你确实要删除该行卫生材料吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim sngLeft As Single, sngTop As Single
    Dim RecReturn As Recordset
    Dim strUnit As String
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    On Error GoTo ErrHandle
    If mshBill.Col = mHeadCol.材料 Then
        Set RecReturn = Frm材料选择器.ShowMe(Me, 1, , mlng库房id, , , , , , , , , , , , , , mstrPrivs, , False)
        If RecReturn.RecordCount > 0 Then
            mblnChange = True
            RecReturn.MoveFirst
            
            If mintUnit = 0 Then
                strUnit = "散装单位"
            Else
                strUnit = "包装单位"
            End If
            
            For i = 1 To RecReturn.RecordCount
                If SetStuffRows(RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                            IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                            Switch(strUnit = "散装单位", RecReturn!散装单位, strUnit = "包装单位", RecReturn!包装单位), RecReturn!指导批发价, _
                            Switch(strUnit = "散装单位", 1, strUnit = "包装单位", RecReturn!换算系数)) Then
                    
                    If mshBill.Row = mshBill.Rows - 1 Then mshBill.Rows = mshBill.Rows + 1 '只有当前行是最后一行时才新增行
                    mshBill.Row = mshBill.Row + 1
                End If
            
                RecReturn.MoveNext
            Next
            
            mshBill.Row = int点击行
            
            If mstr重复卫材 <> "" Then
                MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                mstr重复卫材 = ""
            End If
            
'            If RecReturn.RecordCount = 1 Then
'                If mintUnit = 0 Then
'                    strUnit = "散装单位"
'                Else
'                    strUnit = "包装单位"
'                End If
'                SetStuffRows RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                            IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                            Switch(strUnit = "散装单位", RecReturn!散装单位, strUnit = "包装单位", RecReturn!包装单位), RecReturn!指导批发价, _
'                            Switch(strUnit = "散装单位", 1, strUnit = "包装单位", RecReturn!换算系数)
'            End If
            RecReturn.Close
        End If
    ElseIf mshBill.Col = mHeadCol.上次供应商 Then
        '药品供应商的选择
        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
        If sngLeft + Msf供应商选择.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf供应商选择.Width - 100

        Set RecReturn = New ADODB.Recordset
        gstrSQL = "Select ID,编码,名称,简码 From 供应商 " & _
                  "Where 末级=1 And (站点=[1] or 站点 is null) And (substr(类型,5,1)=1  Or Nvl(末级,0)=0) " & _
                  "  And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) Order By 编码 "
        Set RecReturn = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "-供应商", gstrNodeNo)
        If RecReturn.RecordCount = 0 Then
            MsgBox "请先初始化卫生材料供应商！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        With Msf供应商选择
            .Clear
            Set .DataSource = RecReturn
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800

            .Row = 1
            .ColSel = .Cols - 1
        End With
        With Msf供应商选择
            .Left = sngLeft
            .Top = sngTop
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    ElseIf mshBill.Col = mHeadCol.产地 Then
        '生成商的选择
        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
        If sngLeft + msh生产商.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - msh生产商.Width - 100

        Set RecReturn = New ADODB.Recordset
        gstrSQL = "Select 编码,名称,简码,生产企业许可证,生产企业许可证效期 From 材料生产商 Order By 编码 "
        zlDatabase.OpenRecordset RecReturn, gstrSQL, "读取卫生生产商"
        If RecReturn.RecordCount = 0 Then
            MsgBox "请先初始化卫生材料供应商！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        With msh生产商
            .Clear
            Set .DataSource = RecReturn
            .ColWidth(0) = 800
            .ColWidth(1) = 2000
            .ColWidth(2) = 800
            .ColWidth(3) = 1000
            .ColWidth(4) = 1000

            .Row = 1
            .ColSel = .Cols - 1
        End With
        With msh生产商
            .Left = sngLeft
            .Top = sngTop
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_DblClick(Cancel As Boolean)
    Dim blnAllowChange As Boolean
    Dim lngColor As Long
    Dim i As Integer
    
    With mshBill
        If mblnFirstCheck = False Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .Col <> mHeadCol.校验 Then Exit Sub
        If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
        
        i = .ColData(mHeadCol.材料)
        .ColData(mHeadCol.材料) = 0
        .Col = mHeadCol.材料
        lngColor = .MsfObj.CellForeColor
        If lngColor = vbRed Then
            blnAllowChange = True
        End If
        .ColData(mHeadCol.材料) = i
        
        i = .ColData(mHeadCol.产地)
        .ColData(mHeadCol.产地) = 0
        .Col = mHeadCol.产地
        lngColor = .MsfObj.CellForeColor
        If lngColor = vbRed Then
            blnAllowChange = True
        End If
        .ColData(mHeadCol.产地) = i
        
        i = .ColData(mHeadCol.上次供应商)
        .ColData(mHeadCol.上次供应商) = 0
        .Col = mHeadCol.上次供应商
        lngColor = .MsfObj.CellForeColor
        If lngColor = vbRed Then
           blnAllowChange = True
        End If
        .ColData(mHeadCol.上次供应商) = i
        
        .Col = mHeadCol.校验
        If blnAllowChange = True Then
            If .TextMatrix(.Row, .Col) = "√" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "√"
            End If
        End If
    End With
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mHeadCol.计划数量 Or .Col = mHeadCol.单价 Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mHeadCol.计划数量
                    intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.数量小数, g_小数位数.obj_散装小数.数量小数)
                Case mHeadCol.单价
                    intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.零售价小数, g_小数位数.obj_散装小数.零售价小数)
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                KeyAscii = 0
                Exit Sub
            End If
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If

        Select Case .Col
            Case mHeadCol.材料
                .TxtCheck = False
                .MaxLength = 80
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
            Case mHeadCol.产地
                .TxtCheck = False
                .MaxLength = 40
            Case mHeadCol.上次供应商
                .MaxLength = 40
                .TxtCheck = False
            Case mHeadCol.计划数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mHeadCol.单价
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"

        End Select
        
        Call 显示库存
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsStuff As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    
    Dim rsTemp As Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        If .Col = mHeadCol.材料 Then
            .Text = UCase(Trim(.Text))
        Else
            .Text = Trim(.Text)
        End If
        strKey = .Text

        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col

            Case mHeadCol.材料
                If strKey <> "" Then

                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If

                    Set rsTemp = FrmMulitSel.ShowSelect(Me, 1, , mlng库房id, , strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, , , , , , , , , , , , mstrPrivs, , False)
                    
                    If rsTemp.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    If mintUnit = 0 Then
                        strUnit = "散装单位"
                    Else
                        strUnit = "包装单位"
                    End If
                
                    rsTemp.MoveFirst
                    For i = 1 To rsTemp.RecordCount
                        If SetStuffRows(rsTemp!材料ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, _
                            IIf(IsNull(rsTemp!规格), "", rsTemp!规格), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), _
                            Switch(strUnit = "散装单位", rsTemp!散装单位, strUnit = "包装单位", rsTemp!包装单位), rsTemp!指导批发价, _
                            Switch(strUnit = "散装单位", 1, strUnit = "包装单位", rsTemp!换算系数)) Then
                            
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        rsTemp.MoveNext
                    Next
                    
                    mshBill.Row = int点击行
                    
                    If mstr重复卫材 <> "" Then
                        MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                        mstr重复卫材 = ""
                    End If
                    
'                    If rsTemp.RecordCount = 1 Then
'                        If mintUnit = 0 Then
'                            strUnit = "散装单位"
'                        Else
'                            strUnit = "包装单位"
'                        End If
'                        If SetStuffRows(rsTemp!材料ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, _
'                            IIf(IsNull(rsTemp!规格), "", rsTemp!规格), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), _
'                            Switch(strUnit = "散装单位", rsTemp!散装单位, strUnit = "包装单位", rsTemp!包装单位), rsTemp!指导批发价, _
'                            Switch(strUnit = "散装单位", 1, strUnit = "包装单位", rsTemp!换算系数)) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'
'                        Cancel = True
'                    End If
                End If
            Case mHeadCol.计划数量
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，计划数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "数量必须在(0~99999999)内,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
'                    If .TxtVisible = True Then
'                        .TextMatrix(.Row, mHeadCol.计划数量) = ""
'                    End If
'                    .Col = mHeadCol.材料
'                    If .Row < .Rows - 1 Then
'                        .Row = .Row + 1
'                    Else
'                        If .TextMatrix(.Row, 0) <> "" Then
'                            .Rows = .Rows + 1
'                            .Row = .Row + 1
'                        End If
'                    End If
                    Cancel = True

                    Exit Sub
                End If


                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    If .TextMatrix(.Row, mHeadCol.单价) <> "" Then
                        .TextMatrix(.Row, mHeadCol.金额) = Format(.TextMatrix(.Row, mHeadCol.单价) * strKey, mFMT.FM_金额)
                    End If

                End If
                Call 显示合计金额
            Case mHeadCol.单价
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "单价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "单价必须在(0~99999999)内,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mHeadCol.单价) = " "
                        .Text = " "
                    End If
'                    .Col = mHeadCol.材料
'                    If .Row < .Rows - 1 Then
'                        .Row = .Row + 1
'                    Else
'                        If .TextMatrix(.Row, 0) <> "" Then
'                            .Rows = .Rows + 1
'                            .Row = .Row + 1
'                        End If
'                    End If
'                    .TextMatrix(.Row, mHeadCol.金额) = format(Val(.TextMatrix(.Row, mHeadCol.单价)) * Val(.TextMatrix(.Row, mHeadCol.计划数量)), mFMT.FM_成本价)
                                 
'                    Cancel = True
'                    Exit Sub
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.单价) = strKey
                End If
                .TextMatrix(.Row, mHeadCol.金额) = Format(Val(.TextMatrix(.Row, mHeadCol.单价)) * Val(.TextMatrix(.Row, mHeadCol.计划数量)), mFMT.FM_金额)
                Call 显示合计金额
                
            Case mHeadCol.上次供应商
                If .TxtVisible = False Then Exit Sub
                If strKey = "" And .TextMatrix(.Row, mHeadCol.上次供应商) = "" Then
                    strKey = " "
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.上次供应商) = strKey
                Else
                    If StrIsValid(strKey, 40) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = UCase(strKey)
                    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
                    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngLeft + Msf供应商选择.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf供应商选择.Width - 100
            
                    Set rsTemp = New ADODB.Recordset
                    gstrSQL = "" & _
                        "   Select ID,编码,名称,简码 " & _
                        "   From 供应商 " & _
                        "   Where 末级=1 And (站点=[2] or 站点 is null) And (substr(类型,5,1)=1 Or Nvl(末级,0)=0) " & _
                        "           And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                        "           And (upper(编码) Like [1] Or Upper(名称) Like [1] Or Upper(简码) Like [1])" & _
                        "   Order By 编码 "
                    
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取卫生材料供应商", strKey & "%", gstrNodeNo)
                    
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "没有找到符合条件的供应商！", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    ElseIf rsTemp.RecordCount = 1 Then
                        .Text = rsTemp!名称
                        Exit Sub
                    End If
                    
                    With Msf供应商选择
                        .Clear
                        Set .DataSource = rsTemp
                        .ColWidth(0) = 0
                        .ColWidth(1) = 800
                        .ColWidth(2) = 3000
                        .ColWidth(3) = 800
            
                        .Row = 1
                        .ColSel = .Cols - 1
                    End With
                    With Msf供应商选择
                        .Left = sngLeft
                        .Top = sngTop
                        .Visible = True
                        .ZOrder 0
                        .SetFocus
                    End With
                    Cancel = True
                End If
            Case mHeadCol.产地
'                If .TxtVisible = False Then Exit Sub
                If strKey = "" And .TextMatrix(.Row, mHeadCol.产地) = "" Then
                    strKey = " "
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.产地) = strKey
                Else
                    If strKey <> "" Then
                        If StrIsValid(strKey, 40) = False Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        strKey = UCase(strKey)
                        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
                        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                        If sngLeft + msh生产商.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - msh生产商.Width - 100
                
                        Set rsTemp = New ADODB.Recordset
                        gstrSQL = "" & _
                            "   Select 编码,名称,简码,生产企业许可证,生产企业许可证效期 " & _
                            "   From 材料生产商 " & _
                            "   Where (upper(编码) Like [1] Or Upper(名称) Like [1] Or Upper(简码) Like [1])" & _
                            "   Order By 编码 "
                        
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取材料生产商", strKey & "%")
                        
                        If rsTemp.RecordCount = 0 Then
                            MsgBox "没有找到符合条件的生产商！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        ElseIf rsTemp.RecordCount = 1 Then
                            .Text = rsTemp!名称
                            Exit Sub
                        End If
                        
                        With msh生产商
                            .Clear
                            Set .DataSource = rsTemp
                            .ColWidth(0) = 800
                            .ColWidth(1) = 2000
                            .ColWidth(2) = 800
                            .ColWidth(3) = 1000
                            .ColWidth(4) = 1000
                            
                            .Row = 1
                            .ColSel = .Cols - 1
                        End With
                        With msh生产商
                            .Left = sngLeft
                            .Top = sngTop
                            .Visible = True
                            .ZOrder 0
                            .SetFocus
                        End With
                        Cancel = True
                    End If
                End If
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh生产商_DblClick()
    Dim blnCancel As Boolean
    With mshBill
        .Text = msh生产商.TextMatrix(msh生产商.Row, 1)
        .TextMatrix(.Row, mHeadCol.产地) = msh生产商.TextMatrix(msh生产商.Row, 1)
    End With
    msh生产商.Visible = False
    mshBill.SetFocus
    Call SendKeys("{ENTER}")
End Sub


Private Sub msh生产商_GotFocus()
    If msh生产商.Rows - 1 = 1 Then Call msh生产商_DblClick
End Sub

Private Sub msh生产商_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call msh生产商_DblClick
    End If
End Sub

Private Sub msh生产商_LostFocus()
    msh生产商.ZOrder 1
    msh生产商.Visible = False
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer

    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据

            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > 40 Then
                MsgBox "摘要超长,最多能输入20个汉字或40个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If

            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mHeadCol.材料)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mHeadCol.计划数量))) <> "" Then
                        If Not IsNumeric(.TextMatrix(intLop, mHeadCol.计划数量)) Then
                            MsgBox "第" & intLop & "行卫生材料的计划数量不为数字型，请检查！", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mHeadCol.计划数量
                            Exit Function
                        End If

                    End If
                    
                    If Val(.TextMatrix(intLop, mHeadCol.计划数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫生材料的计划数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.计划数量
                        Exit Function
                    End If

                    If Val(.TextMatrix(intLop, mHeadCol.单价)) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫生材料的单价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.单价
                        Exit Function
                    End If

                    If Val(.TextMatrix(intLop, mHeadCol.金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫生材料的金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.计划数量
                        Exit Function
                    End If
                    
                    If Trim(.TextMatrix(intLop, mHeadCol.上次供应商)) = "" Then
                        MsgBox "第" & intLop & "行卫生材料未选择供应商，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.上次供应商
                        Exit Function
                    End If
                
                End If
            Next
        Else
            Exit Function
        End If
    End With

    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim lng序号 As Long
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim 计划类型_IN As Integer
    Dim 期间_IN As String
    Dim 库房ID_IN As Long
    Dim 编制方法_IN As Integer
    Dim 编制人_IN As String
    Dim 编制日期_IN As String
    Dim 编制说明_IN As String

    Dim 材料ID_IN As Long
    Dim 计划数量_IN As Double
    Dim 单价_IN As Double
    Dim 金额_IN As Double
    Dim 前期数量_IN As Double
    Dim 上期数量_IN As Double
    Dim 库存数量_IN As Double
    Dim 上期销量_IN As Double
    Dim 本期销量_IN As Double
    Dim 上次供应商_IN As String
    Dim 上次生产商_IN As String
    Dim intRow As Integer
    Dim cllTemp As New Collection
    SaveCard = False
    With mshBill
        ID_IN = sys.NextId("材料采购计划")
        NO_IN = Trim(txtNO)
        
        If NO_IN = "" Then NO_IN = sys.GetNextNo(77, mlng库房id)
        If IsNull(NO_IN) Then Exit Function
        Me.txtNO.Tag = NO_IN
        
        计划类型_IN = mint计划类型
        编制方法_IN = mint编制方法
        库房ID_IN = mlng库房id
        编制人_IN = gstrUserName
        编制日期_IN = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        编制说明_IN = Trim(txt摘要.Text)
        期间_IN = mstr期间

        If mint编辑状态 = 2 Then        '修改
            gstrSQL = "zl_材料计划管理_DELETE('" & mlng计划ID & "')"
            cllTemp.Add gstrSQL
        End If
        'Zl_材料计划管理主表_Insert
        gstrSQL = "Zl_材料计划管理主表_Insert("
        '  Id_In       In 材料采购计划.ID%Type,
        gstrSQL = gstrSQL & "" & ID_IN & ","
        '  单据_In     In 材料采购计划.单据%Type,
        gstrSQL = gstrSQL & "" & 0 & ","
        '  No_In       In 材料采购计划.NO%Type,
        gstrSQL = gstrSQL & "'" & NO_IN & "',"
        '  计划类型_In In 材料采购计划.计划类型%Type,
        gstrSQL = gstrSQL & "" & 计划类型_IN & ","
        '  期间_In     In 材料采购计划.期间%Type,
        gstrSQL = gstrSQL & "'" & 期间_IN & "',"
        '  库房id_In   In 材料采购计划.库房id%Type,
        gstrSQL = gstrSQL & "" & IIf(库房ID_IN = 0, "NULL", 库房ID_IN) & ","
        '  部门id_In   In 材料采购计划.部门id%Type,
        gstrSQL = gstrSQL & "NULL,"
        '  编制方法_In In 材料采购计划.编制方法%Type,
        gstrSQL = gstrSQL & "" & 编制方法_IN & ","
        '  编制人_In   In 材料采购计划.编制人%Type,
        gstrSQL = gstrSQL & "'" & 编制人_IN & "',"
        '  编制日期_In In 材料采购计划.编制日期%Type,
        gstrSQL = gstrSQL & "to_date('" & 编制日期_IN & "','yyyy-mm-dd HH24:MI:SS'),"
        '  编制说明_In In 材料采购计划.编制说明%Type := Null
        gstrSQL = gstrSQL & "'" & 编制说明_IN & "')"
        cllTemp.Add gstrSQL
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng序号 = .TextMatrix(intRow, mHeadCol.序号)
                材料ID_IN = .TextMatrix(intRow, 0)
                单价_IN = Round(Val(.TextMatrix(intRow, mHeadCol.单价)) / Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.成本价小数)
                金额_IN = Round(Val(.TextMatrix(intRow, mHeadCol.金额)), g_小数位数.obj_最大小数.金额小数)
                前期数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.前期数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.数量小数)
                上期数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.上期数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.数量小数)
                库存数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.库存数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.数量小数)
                计划数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.计划数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.数量小数)
                上期销量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.上期销量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.数量小数)
                本期销量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.本期销量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.数量小数)
                上次供应商_IN = .TextMatrix(intRow, mHeadCol.上次供应商)
                上次生产商_IN = .TextMatrix(intRow, mHeadCol.产地)
                'zl_药品计划管理次表_INSERT( /*计划ID_IN*/, /*材料ID_IN*/,/请购数量_IN /*计划数量_IN*/,
                    '/*单价_IN*/, /*金额_IN*/, /*前期数量_IN*/, /*上期数量_IN*/, /*库存数量_IN*/,
                    '/*上次供应商_IN*/, /*上次生产商_IN*/ );

                gstrSQL = "zl_材料计划管理次表_INSERT(" & ID_IN & "," & 材料ID_IN & "," & lng序号 & ",0," & 计划数量_IN _
                    & "," & 单价_IN & "," & 金额_IN & "," & 前期数量_IN & "," & 上期数量_IN _
                    & "," & 库存数量_IN & ",'" & 上次供应商_IN & "','" & 上次生产商_IN & "'," & 上期销量_IN & "," & 本期销量_IN & ")"
                cllTemp.Add gstrSQL
            End If
        Next
    End With
    On Error GoTo ErrHandle
    ExecuteProcedureArrAy cllTemp, mstrCaption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveCheckCard() As Boolean
    '审核时，保存通过审核的单据（校验列打勾的列）
    Dim lng序号 As Long
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim 计划类型_IN As Integer
    Dim 期间_IN As String
    Dim 库房ID_IN As Long
    Dim 编制方法_IN As Integer
    Dim 编制人_IN As String
    Dim 编制日期_IN As String
    Dim 编制说明_IN As String

    Dim 材料ID_IN As Long
    Dim 计划数量_IN As Double
    Dim 单价_IN As Double
    Dim 金额_IN As Double
    Dim 前期数量_IN As Double
    Dim 上期数量_IN As Double
    Dim 库存数量_IN As Double
    Dim 上期销量_IN As Double
    Dim 本期销量_IN As Double
    Dim 上次供应商_IN As String
    Dim 上次生产商_IN As String
    Dim intRow As Integer
    Dim cllTemp As New Collection
    Dim blnNoRecord As Boolean
    
    blnNoRecord = True
    
    SaveCheckCard = False
    
    With mshBill
        ID_IN = sys.NextId("材料采购计划")
        NO_IN = Trim(txtNO)
        
        If NO_IN = "" Then NO_IN = sys.GetNextNo(77, mlng库房id)
        If IsNull(NO_IN) Then Exit Function
        Me.txtNO.Tag = NO_IN
        
        计划类型_IN = mint计划类型
        编制方法_IN = mint编制方法
        库房ID_IN = mlng库房id
        编制人_IN = IIf(Txt填制人.Caption <> "", Txt填制人, gstrUserName)
        编制日期_IN = IIf(Txt填制日期.Caption <> "", Format(Txt填制日期.Caption, "yyyy-mm-dd hh:mm:ss"), Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss"))
        编制说明_IN = Trim(txt摘要.Text)
        期间_IN = mstr期间

        '删除原来的单据
        gstrSQL = "zl_材料计划管理_DELETE('" & mlng计划ID & "')"
        cllTemp.Add gstrSQL
        
        'Zl_材料计划管理主表_Insert
        gstrSQL = "Zl_材料计划管理主表_Insert("
        '  Id_In       In 材料采购计划.ID%Type,
        gstrSQL = gstrSQL & "" & ID_IN & ","
        '  单据_In     In 材料采购计划.单据%Type,
        gstrSQL = gstrSQL & "" & 0 & ","
        '  No_In       In 材料采购计划.NO%Type,
        gstrSQL = gstrSQL & "'" & NO_IN & "',"
        '  计划类型_In In 材料采购计划.计划类型%Type,
        gstrSQL = gstrSQL & "" & 计划类型_IN & ","
        '  期间_In     In 材料采购计划.期间%Type,
        gstrSQL = gstrSQL & "'" & 期间_IN & "',"
        '  库房id_In   In 材料采购计划.库房id%Type,
        gstrSQL = gstrSQL & "" & IIf(库房ID_IN = 0, "NULL", 库房ID_IN) & ","
        '  部门id_In   In 材料采购计划.部门id%Type,
        gstrSQL = gstrSQL & "NULL,"
        '  编制方法_In In 材料采购计划.编制方法%Type,
        gstrSQL = gstrSQL & "" & 编制方法_IN & ","
        '  编制人_In   In 材料采购计划.编制人%Type,
        gstrSQL = gstrSQL & "'" & 编制人_IN & "',"
        '  编制日期_In In 材料采购计划.编制日期%Type,
        gstrSQL = gstrSQL & "to_date('" & 编制日期_IN & "','yyyy-mm-dd HH24:MI:SS'),"
        '  编制说明_In In 材料采购计划.编制说明%Type := Null
        gstrSQL = gstrSQL & "'" & 编制说明_IN & "')"
        cllTemp.Add gstrSQL
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And .TextMatrix(intRow, mHeadCol.校验) = "√" Then
                lng序号 = lng序号 + 1
                材料ID_IN = .TextMatrix(intRow, 0)
                单价_IN = Round(Val(.TextMatrix(intRow, mHeadCol.单价)) / Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.成本价小数)
                金额_IN = Round(Val(.TextMatrix(intRow, mHeadCol.金额)), g_小数位数.obj_散装小数.金额小数)
                前期数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.前期数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.数量小数)
                上期数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.上期数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.数量小数)
                库存数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.库存数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.数量小数)
                计划数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.计划数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.数量小数)
                上期销量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.上期销量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.数量小数)
                本期销量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.本期销量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.数量小数)
                上次供应商_IN = .TextMatrix(intRow, mHeadCol.上次供应商)
                上次生产商_IN = .TextMatrix(intRow, mHeadCol.产地)
                'zl_药品计划管理次表_INSERT( /*计划ID_IN*/, /*材料ID_IN*/,/请购数量_IN /*计划数量_IN*/,
                    '/*单价_IN*/, /*金额_IN*/, /*前期数量_IN*/, /*上期数量_IN*/, /*库存数量_IN*/,
                    '/*上次供应商_IN*/, /*上次生产商_IN*/ );

                gstrSQL = "zl_材料计划管理次表_INSERT(" & ID_IN & "," & 材料ID_IN & "," & lng序号 & ",0," & 计划数量_IN _
                    & "," & 单价_IN & "," & 金额_IN & "," & 前期数量_IN & "," & 上期数量_IN _
                    & "," & 库存数量_IN & ",'" & 上次供应商_IN & "','" & 上次生产商_IN & "'," & 上期销量_IN & "," & 本期销量_IN & ")"
                cllTemp.Add gstrSQL
                
                blnNoRecord = False
            End If
        Next
    End With
    
    If blnNoRecord = True Then
        MsgBox "请选择允许审核通过的记录。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mlng计划ID = ID_IN

    On Error GoTo ErrHandle
    
    ExecuteProcedureArrAy cllTemp, mstrCaption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCheckCard = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub 显示合计金额()
    Dim Dbl金额 As Double
    Dim intLop As Integer

    Dbl金额 = 0

    With mshBill
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                Dbl金额 = Dbl金额 + Val(.TextMatrix(intLop, mHeadCol.金额))
            End If
        Next
    End With

    lblPurchasePrice.Caption = "金额合计：" & Format(Dbl金额, mFMT.FM_金额)
End Sub

Private Sub txt摘要_Change()
    mblnChange = True
End Sub

Private Sub txt摘要_GotFocus()
    OS.OpenIme (True)
    With txt摘要
        .SelStart = 0
        .SelLength = Len(txt摘要.Text)
    End With
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt摘要_LostFocus()
    OS.OpenIme False
End Sub

Private Function SetStuffRows(ByVal lng材料ID As Long, ByVal str材料 As String, _
        ByVal str规格 As String, ByVal str产地 As String, ByVal str单位 As String, _
        ByVal dbl指导批发价 As Double, ByVal dbl比例系数 As Double) As Boolean
    Dim rsData As New Recordset
    Dim intCount As Integer
    Dim intRow As Integer
    Dim intCol As Integer

    Dim lng批次 As Long
    Dim dbl库存数量 As Double
    Dim dbl成本价 As Double

    On Error GoTo errH
    SetStuffRows = False

    With mshBill
        intRow = .Row
        For intCount = 1 To .Rows - 1
            If intCount <> intRow And .TextMatrix(intCount, 0) <> "" Then
                If .TextMatrix(intCount, 0) = lng材料ID Then
                    If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & str材料 & "，"  '最多记录三个重复的卫材
                    'MsgBox "对不起，卫生材料【" & str材料 & "】已有了，不能再输！", vbOKOnly + vbExclamation, gstrSysName
                    Exit Function
                End If
            End If
        Next

        For intCol = 0 To .Cols - 1
            .TextMatrix(intRow, intCol) = ""
        Next
    End With

    With mshBill
        .TextMatrix(.Row, mHeadCol.序号) = .Row
        .TextMatrix(.Row, mHeadCol.产地) = str产地
        .TextMatrix(.Row, 0) = lng材料ID
        .TextMatrix(.Row, mHeadCol.比例系数) = dbl比例系数
        
        '取平均成本价（如果没有设置，则取指导批发价）
        gstrSQL = "Select 成本价,指导批发价,招标材料 From  材料特性 Where 材料ID=[1]"
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "取成本价", lng材料ID)
        
        dbl成本价 = zlStr.Nvl(rsData!成本价, 0)
        If dbl成本价 = 0 Then dbl成本价 = zlStr.Nvl(rsData!指导批发价, 0)
        .TextMatrix(.Row, mHeadCol.中标材料) = IIf(Val(zlStr.Nvl(rsData!招标材料)) = "1", "√", "")
        
        gstrSQL = "" & _
            " SELECT MIN (B.名称) AS 供应商, MIN (上次产地) AS 上次产地,SUM(实际数量) AS 库存数量 " & _
            " FROM 药品库存 A, (SELECT ID,名称 FROM 供应商 WHERE SUBSTR(类型,5,1)=1) B  " & _
            " WHERE A.性质=1 AND A.上次供应商ID = B.ID(+) And A.药品ID=[1] " & _
            IIf(mlng库房id = 0, "", " AND A.库房ID=[2]")
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "取上次供应商及产地信息", lng材料ID, mlng库房id)
        
        If zlStr.Nvl(rsData!库存数量, 0) = 0 And zlStr.Nvl(rsData!供应商) = "" And zlStr.Nvl(rsData!上次产地) = "" Then
            gstrSQL = "Select c.名称 As 供应商, Decode(a.上次产地, Null, b.产地, a.上次产地) As 上次产地, 0 As 库存数量" & _
                       " From 材料特性 A, 收费项目目录 B, (Select ID, 名称 From 供应商 Where Substr(类型, 5, 1) = 1 And (站点 = [3] Or 站点 Is Null)) C" & _
                       " Where a.材料id = b.Id And a.上次供应商id = c.Id And a.材料id = [1]"
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "取上次供应商及产地信息", lng材料ID, mlng库房id, gstrNodeNo)
        End If
            
        If Not rsData.EOF Then
            .TextMatrix(.Row, mHeadCol.库存数量) = Format(IIf(IsNull(rsData!库存数量), 0, rsData!库存数量) / dbl比例系数, mFMT.FM_数量)
            .TextMatrix(.Row, mHeadCol.上次供应商) = IIf(IsNull(rsData!供应商), "", rsData!供应商)
            .TextMatrix(.Row, mHeadCol.产地) = IIf(IsNull(rsData!上次产地), str产地, rsData!上次产地)
            SetNumer lng材料ID, mlng库房id, .TextMatrix(.Row, mHeadCol.库存数量), .Row, mint计划类型, mint编制方法
        End If
        .TextMatrix(.Row, mHeadCol.材料) = str材料
        .TextMatrix(.Row, mHeadCol.规格) = str规格
        .TextMatrix(.Row, mHeadCol.单位) = str单位
        .TextMatrix(.Row, mHeadCol.单价) = Format(dbl成本价 * dbl比例系数, mFMT.FM_成本价)
        
        
        gstrSQL = "" & _
            "   Select sum(nvl(上限,0)) 上限,sum(nvl(下限,0)) 下限 " & _
            "   From 材料储备限额  " & _
            "   where 材料ID=[1] " & IIf(mlng库房id = 0, "", " and 库房ID=[2]") & _
            "   Group By 材料ID"
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "取上下限", lng材料ID, mlng库房id)
        If Not rsData.EOF Then
            .TextMatrix(.Row, mHeadCol.存储上限) = Format(Val(zlStr.Nvl(rsData!上限)) / dbl比例系数, mFMT.FM_数量)
            .TextMatrix(.Row, mHeadCol.存储下限) = Format(Val(zlStr.Nvl(rsData!下限)) / dbl比例系数, mFMT.FM_数量)
        End If
        
    End With
    rsData.Close
    SetStuffRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Private Sub 显示库存()
    Dim rsData As ADODB.Recordset
    Dim lng材料ID As Long
    Dim str单位 As String
    Dim dbl包装 As Double
    Dim strTmp As String
    
    If mblnStart = False Then Exit Sub
    
    On Error GoTo ErrHandle
    Me.stbThis.Panels(2).Text = ""
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    
    lng材料ID = Val(mshBill.TextMatrix(mshBill.Row, 0))
    str单位 = mshBill.TextMatrix(mshBill.Row, mHeadCol.单位)
    dbl包装 = Val(mshBill.TextMatrix(mshBill.Row, mHeadCol.比例系数))
    
    gstrSQL = "Select B.名称, A.药品id, Nvl(Sum(A.实际数量),0) As 实际数量 " & _
        " From 药品库存 A, 部门表 B " & _
        " Where A.性质 = 1 And A.库房id + 0 = B.ID And A.药品id = [1] " & _
        " Group By B.名称, A.药品id " & _
        " Order By B.名称"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "显示库存", lng材料ID)
    
    Do While Not rsData.EOF
        strTmp = IIf(strTmp = "", "", strTmp & ";") & rsData!名称 & zlStr.FormatEx(rsData!实际数量 / dbl包装, 2, , True) & str单位
        rsData.MoveNext
    Loop
    
    If strTmp <> "" Then Me.stbThis.Panels(2).Text = "该药品当前库存：" & strTmp
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'按编码，名称，别名查找某一列
Private Function FindData(ByVal mshBill As BillEdit, ByVal int比较列 As Integer, _
    ByVal str比较值 As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    Dim strKey As String
    FindData = True
    
    On Error GoTo ErrHandle
    With mshBill
        If .Rows = 2 Then Exit Function
        If str比较值 = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                If InStr(1, UCase(strCode), UCase(str比较值)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int比较列
                    .MsfObj.TopRow = .Row
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = " SELECT DISTINCT b.编码 " & _
                  " FROM " & _
                  "    (SELECT DISTINCT A.收费细目id " & _
                  "    FROM 收费项目别名 A" & _
                  "    Where A.简码 LIKE [1]) a," & _
                  " 收费项目目录 B " & _
                  " Where a.收费细目id = b.ID"
        
        strKey = IIf(gstrMatchMethod = "0", "%", "") & str比较值 & "%"
        Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strKey)
                  
        If rsCode.EOF Then
            FindData = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!编码)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int比较列
                        .MsfObj.TopRow = .Row
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    FindData = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

